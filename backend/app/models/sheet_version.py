import uuid
from datetime import datetime, timezone, timedelta
from sqlalchemy import Column, String, DateTime, Integer, Boolean, JSON
from app.database import Base

IST_OFFSET = timezone(timedelta(hours=5, minutes=30))

def now_ist() -> datetime:
    return datetime.now(IST_OFFSET).replace(tzinfo=None)

class SheetVersion(Base):
    __tablename__ = "sheet_versions"
    
    id = Column(String(36), primary_key=True, default=lambda: str(uuid.uuid4()))
    document_id = Column(String(36), nullable=False, index=True)
    version_name = Column(String(255), nullable=True)
    created_by_user_id = Column(Integer, nullable=False)
    contributors = Column(JSON, nullable=False)
    is_named = Column(Boolean, default=False)
    sheet_snapshot_url = Column(String(500), nullable=False)
    base_version_id = Column(String(36), nullable=True)
    created_at = Column(DateTime, default=now_ist, index=True)
