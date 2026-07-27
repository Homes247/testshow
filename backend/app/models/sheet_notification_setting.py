from datetime import datetime
from sqlalchemy import Integer, String, DateTime, SmallInteger, func, Text
from sqlalchemy.orm import Mapped, mapped_column
from app.database import Base


class SheetNotificationSetting(Base):
    """
    Stores per-user email notification preferences for a given sheet.
    When a user opts in, notifications are delivered as internal Vmail messages.
    """
    __tablename__ = "sheet_notification_settings"
    __table_args__ = {"extend_existing": True}

    id:            Mapped[int]      = mapped_column(Integer,      primary_key=True, autoincrement=True)
    document_id:   Mapped[str]      = mapped_column(String(64),   nullable=False, index=True)
    user_id:       Mapped[int]      = mapped_column(Integer,      nullable=False, index=True)
    notify_email:  Mapped[str]      = mapped_column(String(255),  nullable=True)
    on_edit:       Mapped[int]      = mapped_column(SmallInteger, default=1)
    on_comment:    Mapped[int]      = mapped_column(SmallInteger, default=1)
    created_at:    Mapped[datetime] = mapped_column(DateTime,     server_default=func.now())
    updated_at:    Mapped[datetime] = mapped_column(DateTime,     server_default=func.now(), onupdate=func.now())
