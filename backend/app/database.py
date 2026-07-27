import os
from dotenv import load_dotenv
from sqlalchemy.ext.asyncio import create_async_engine, AsyncSession, async_sessionmaker
from sqlalchemy.orm import DeclarativeBase

load_dotenv(override=True)

# ── Connects to the VMAIL database ──────────────────────────
# GDoc has no separate DB — all tables (documents, document_shares)
# live inside the vmail DB alongside users, mails, etc.
DB_HOST = os.getenv("DB_HOST", "74.225.249.46")
DB_PORT = os.getenv("DB_PORT", "3307")
DB_NAME = os.getenv("DB_NAME", "vmail")
DB_USER = os.getenv("DB_USER", "chat_bot")
DB_PASS = os.getenv("DB_PASSWORD", "password")

DATABASE_URL = (
    f"mysql+aiomysql://{DB_USER}:{DB_PASS}@{DB_HOST}:{DB_PORT}/{DB_NAME}?charset=utf8mb4"
)

engine = create_async_engine(
    DATABASE_URL,
    echo=False,
    pool_size=10,
    max_overflow=20,
    pool_recycle=300,
    pool_timeout=10,
    pool_pre_ping=True,   # Enabled to prevent 'Lost connection' OperationalErrors crashing WebSockets
    connect_args={
        "connect_timeout": 10
    }
)

SessionLocal = async_sessionmaker(engine, class_=AsyncSession, expire_on_commit=False)

class Base(DeclarativeBase):
    pass

async def get_db():
    async with SessionLocal() as session:
        try:
            yield session
            await session.commit()
        except Exception:
            await session.rollback()
            raise