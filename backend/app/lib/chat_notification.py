"""
chat_notification.py
─────────────────────
Sends internal chat messages directly to the shared Vmail MySQL database.
This is used to send chat links when a sheet is shared.
"""

import os
import logging
from typing import Optional
from datetime import datetime
from zoneinfo import ZoneInfo

from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession

log = logging.getLogger("chat_notification")

IST = ZoneInfo("Asia/Kolkata")

def now_ist() -> datetime:
    return datetime.now(IST).replace(tzinfo=None)

async def _get_or_create_conversation(db: AsyncSession, user1_id: int, user2_id: int) -> int:
    uid1, uid2 = sorted([user1_id, user2_id])
    result = await db.execute(
        text("SELECT id FROM chat_conversations WHERE user1_id = :u1 AND user2_id = :u2 LIMIT 1"),
        {"u1": uid1, "u2": uid2}
    )
    row = result.fetchone()
    if row:
        conv_id = row[0]
        # Update timestamp
        await db.execute(
            text("UPDATE chat_conversations SET updated_at = :now WHERE id = :id"),
            {"now": now_ist(), "id": conv_id}
        )
        return conv_id
    
    # Create new conversation
    result = await db.execute(
        text("INSERT INTO chat_conversations (user1_id, user2_id, created_at, updated_at) VALUES (:u1, :u2, :now, :now)"),
        {"u1": uid1, "u2": uid2, "now": now_ist()}
    )
    return result.lastrowid

async def send_chat_message(
    db: AsyncSession,
    *,
    sender_id: int,
    receiver_id: int,
    message: str
) -> bool:
    """
    Sends a direct chat message from sender to receiver.
    Returns True on success.
    """
    try:
        conv_id = await _get_or_create_conversation(db, sender_id, receiver_id)
        
        await db.execute(
            text("""
                INSERT INTO chat_messages 
                    (conversation_id, sender_id, message, is_file, is_read, delete_status, created_at, updated_at) 
                VALUES 
                    (:conv_id, :sender_id, :msg, 0, 0, 0, :now, :now)
            """),
            {
                "conv_id": conv_id,
                "sender_id": sender_id,
                "msg": message,
                "now": now_ist()
            }
        )
        await db.commit()
        return True
    except Exception as exc:
        log.error("Failed to send chat notification: %s", exc, exc_info=True)
        try:
            await db.rollback()
        except Exception:
            pass
        return False
