import asyncio
import os
from dotenv import load_dotenv
from sqlalchemy import text
from sqlalchemy.ext.asyncio import create_async_engine

load_dotenv()
DB_HOST = os.getenv("DB_HOST", "74.225.249.46")
DB_PORT = os.getenv("DB_PORT", "3307")
DB_NAME = os.getenv("DB_NAME", "vmail")
DB_USER = os.getenv("DB_USER", "chat_bot")
DB_PASS = os.getenv("DB_PASSWORD", "password")

DATABASE_URL = f"mysql+aiomysql://{DB_USER}:{DB_PASS}@{DB_HOST}:{DB_PORT}/{DB_NAME}?charset=utf8mb4"

async def main():
    engine = create_async_engine(DATABASE_URL)
    async with engine.begin() as conn:
        res = await conn.execute(text("SELECT COUNT(*) FROM audit_events"))
        print(f"Total audit events: {res.scalar()}")
    await engine.dispose()

if __name__ == "__main__":
    asyncio.run(main())
