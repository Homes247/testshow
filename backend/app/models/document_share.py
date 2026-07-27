import uuid
from datetime import datetime, timezone, timedelta
from sqlalchemy import String, DateTime, Integer
from sqlalchemy.orm import Mapped, mapped_column
from app.database import Base

IST_OFFSET = timezone(timedelta(hours=5, minutes=30))

def now_ist() -> datetime:
    return datetime.now(IST_OFFSET).replace(tzinfo=None)

class DocumentShare(Base):
    __tablename__ = "document_shares"
    __table_args__ = {"extend_existing": True}

    id:             Mapped[str]           = mapped_column(String(36),  primary_key=True, default=lambda: str(uuid.uuid4()))
    document_id:    Mapped[str]           = mapped_column(String(36),  nullable=False)
    user_id:        Mapped[int]           = mapped_column(Integer,     nullable=True)   # FK → vmail users.id (INT), NULL = external share
    external_email: Mapped[str]           = mapped_column(String(200), nullable=True)   # Email for non-vmail users
    permission:     Mapped[str]           = mapped_column(String(10),  default="view")
    created_at:     Mapped[datetime]      = mapped_column(DateTime,    default=now_ist)