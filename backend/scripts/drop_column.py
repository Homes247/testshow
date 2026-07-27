import asyncio
import os
import sys

# Add backend directory to sys.path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app.database import engine
from sqlalchemy import text

async def drop_content_column():
    async with engine.begin() as conn:
        try:
            await conn.execute(text("ALTER TABLE documents DROP COLUMN content"))
            print("Dropped content column successfully")
        except Exception as e:
            print(f"Error dropping content column: {e}")

if __name__ == "__main__":
    asyncio.run(drop_content_column())
