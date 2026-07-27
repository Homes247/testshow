import uuid
from datetime import datetime, timezone, timedelta
from sqlalchemy import String, DateTime, Integer, JSON
from sqlalchemy.orm import Mapped, mapped_column
from app.database import Base

IST_OFFSET = timezone(timedelta(hours=5, minutes=30))

def now_ist() -> datetime:
    return datetime.now(IST_OFFSET).replace(tzinfo=None)

class AuditEvent(Base):
    __tablename__ = "audit_events"

    id:           Mapped[str]      = mapped_column(String(36),  primary_key=True, default=lambda: str(uuid.uuid4()))
    document_id:  Mapped[str]      = mapped_column(String(36),  nullable=False, index=True)
    user_name:    Mapped[str]      = mapped_column(String(100),  nullable=False)
    sheet_id:     Mapped[str]      = mapped_column(String(100),  nullable=False)
    sheet_name:   Mapped[str]      = mapped_column(String(100),  nullable=False)
    action_type:  Mapped[str]      = mapped_column(String(100),  nullable=False)
    target_range: Mapped[str]      = mapped_column(String(100),  nullable=False)
    metadata_json:Mapped[dict]     = mapped_column(JSON,         nullable=True)
    created_at:   Mapped[datetime] = mapped_column(DateTime,     default=now_ist, index=True)
