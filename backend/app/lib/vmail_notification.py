"""
vmail_notification.py
─────────────────────
Sends internal Vmail notifications by writing directly to the shared MySQL DB
that both the office-suite backend and the Vmail backend use.

The office-suite backend already shares the same `vmail` MySQL DB
(see DATABASE_URL in .env), so we simply INSERT into the `mails` and
`mail_recipients` tables using the existing SQLAlchemy session —
no HTTP call needed.

Notification email: a system account (e.g. no-reply@vsnapmail.co.in) sends
the message to the recipient's Vmail inbox. Both sender and recipient must
exist in the `users` table. We use a fixed system sender_id that you can
configure via VMAIL_NOTIFICATION_SENDER_EMAIL env var.
"""

import os
import logging
from typing import Optional
from datetime import datetime

from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession

log = logging.getLogger("vmail_notification")

# ── Config ─────────────────────────────────────────────────────────────────
# Email of the internal system/notification account in Vmail.
# Set VMAIL_NOTIFICATION_SENDER_EMAIL in .env, or it falls back to this.
NOTIFICATION_SENDER_EMAIL = os.getenv(
    "VMAIL_NOTIFICATION_SENDER_EMAIL", "no-reply@vsnapmail.co.in"
)

SHEETS_URL = os.getenv("SHEETS_URL", "https://testsheets.vsnaptechnology.com")


# ── Helpers ─────────────────────────────────────────────────────────────────

async def _get_user_id_by_email(db: AsyncSession, email: str) -> Optional[int]:
    result = await db.execute(
        text("SELECT id FROM users WHERE email = :email LIMIT 1"),
        {"email": email}
    )
    row = result.fetchone()
    return row[0] if row else None


async def _get_organization_id(db: AsyncSession, user_id: int) -> Optional[int]:
    result = await db.execute(
        text("SELECT organization_id FROM users WHERE id = :uid LIMIT 1"),
        {"uid": user_id}
    )
    row = result.fetchone()
    return row[0] if row else None


# ── Main function ────────────────────────────────────────────────────────────

async def send_sheet_notification(
    db: AsyncSession,
    *,
    recipient_email: str,
    document_id: str,
    document_title: str,
    event_type: str,          # "edit" | "comment"
    actor_name: str,
    actor_email: str,
    extra_detail: str = "",
) -> bool:
    """
    Creates an internal Vmail message from the notification sender to the
    recipient for a sheet edit or comment event.

    Returns True on success, False on any failure (non-fatal — caller should
    log but not raise).
    """
    try:
        # 1. Resolve sender (notification system account)
        sender_id = await _get_user_id_by_email(db, NOTIFICATION_SENDER_EMAIL)
        if sender_id is None:
            log.warning(
                "Vmail notification sender '%s' not found in users table — "
                "skipping notification.", NOTIFICATION_SENDER_EMAIL
            )
            return False

        # 2. Resolve recipient
        recipient_id = await _get_user_id_by_email(db, recipient_email)
        if recipient_id is None:
            log.warning(
                "Notification recipient '%s' not found in users table — "
                "skipping.", recipient_email
            )
            return False

        # 3. Build subject & body
        action_label = "edited" if event_type == "edit" else "commented on"
        subject = f"[Sheets] {actor_name} {action_label} '{document_title}'"

        sheet_url = f"{SHEETS_URL}/sheet/{document_id}"
        body = f"""<div style="font-family:Arial,sans-serif;font-size:14px;color:#333;">
<p>Hi,</p>
<p><strong>{actor_name}</strong> ({actor_email}) has {action_label} the sheet
<strong>{document_title}</strong>.</p>
{f'<p><em>{extra_detail}</em></p>' if extra_detail else ''}
<p>
  <a href="{sheet_url}"
     style="background:#10b981;color:#fff;padding:8px 18px;border-radius:4px;
            text-decoration:none;display:inline-block;">
    Open Sheet
  </a>
</p>
<hr style="border:none;border-top:1px solid #e2e8f0;margin:20px 0;">
<p style="font-size:12px;color:#888;">
  You are receiving this because you enabled email notifications for this sheet.<br>
  To unsubscribe, open the sheet → Tools → Email Notification Settings → uncheck options.
</p>
</div>"""

        body_size   = len(body.encode("utf-8"))
        subject_size = len(subject.encode("utf-8"))
        now = datetime.utcnow()

        # 4. Insert mail row
        mail_result = await db.execute(
            text("""
                INSERT INTO mails
                    (sender_id, sender_email, subject, body, is_draft,
                     delivery_status, size_bytes, created_at, updated_at,
                     thread_id)
                VALUES
                    (:sender_id, :sender_email, :subject, :body, 0,
                     'sent', :size_bytes, :now, :now, NULL)
            """),
            {
                "sender_id":    sender_id,
                "sender_email": NOTIFICATION_SENDER_EMAIL,
                "subject":      subject,
                "body":         body,
                "size_bytes":   body_size + subject_size,
                "now":          now,
            }
        )
        mail_id = mail_result.lastrowid

        # Set thread_id = mail_id (self-thread)
        await db.execute(
            text("UPDATE mails SET thread_id = :mid WHERE id = :mid"),
            {"mid": mail_id}
        )

        # 5. Insert mail_recipient row (inbox)
        await db.execute(
            text("""
                INSERT INTO mail_recipients
                    (mail_id, recipient_id, recipient_type, folder,
                     is_read, is_archived, is_flagged)
                VALUES
                    (:mail_id, :recipient_id, 'to', 'inbox', 0, 0, 0)
            """),
            {"mail_id": mail_id, "recipient_id": recipient_id}
        )

        # 6. Insert sender_folder row (sent)
        await db.execute(
            text("""
                INSERT INTO sender_folders
                    (mail_id, user_id, is_archived)
                VALUES
                    (:mail_id, :sender_id, 0)
            """),
            {"mail_id": mail_id, "sender_id": sender_id}
        )

        await db.commit()
        log.info(
            "Vmail notification sent: mail_id=%s recipient=%s event=%s doc=%s",
            mail_id, recipient_email, event_type, document_id
        )
        return True

    except Exception as exc:
        log.error("Failed to send Vmail notification: %s", exc, exc_info=True)
        try:
            await db.rollback()
        except Exception:
            pass
        return False
