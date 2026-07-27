"""
Migration: create sheet_notification_settings table
Run with: python scripts/create_notification_settings_table.py
"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

import asyncio
from sqlalchemy import text
from app.database import engine


CREATE_SQL = """
CREATE TABLE IF NOT EXISTS sheet_notification_settings (
    id            INT            NOT NULL AUTO_INCREMENT PRIMARY KEY,
    document_id   VARCHAR(64)    NOT NULL,
    user_id       INT            NOT NULL,
    notify_email  VARCHAR(255)   NULL,
    on_edit       TINYINT(1)     NOT NULL DEFAULT 1,
    on_comment    TINYINT(1)     NOT NULL DEFAULT 1,
    created_at    DATETIME       NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at    DATETIME       NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    INDEX idx_sns_doc  (document_id),
    INDEX idx_sns_user (user_id),
    UNIQUE KEY uq_sns_doc_user (document_id, user_id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
"""


async def run():
    async with engine.begin() as conn:
        await conn.execute(text(CREATE_SQL))
    print("[OK] sheet_notification_settings table created (or already exists).")


if __name__ == "__main__":
    asyncio.run(run())
