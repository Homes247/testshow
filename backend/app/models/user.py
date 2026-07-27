from datetime import datetime
from sqlalchemy import Integer, String, DateTime, SmallInteger, Float, func
from sqlalchemy.orm import Mapped, mapped_column
from app.database import Base


class User(Base):
    """
    Maps to the existing VMail `users` table.
    Do NOT let SQLAlchemy create or modify this table —
    it is owned by the VMail backend.
    """
    __tablename__ = "users"
    __table_args__ = {"extend_existing": True}

    id:              Mapped[int]      = mapped_column(Integer,      primary_key=True)
    name:            Mapped[str]      = mapped_column(String(100))
    email:           Mapped[str]      = mapped_column(String(100),  unique=True, index=True)
    password_hash:   Mapped[str]      = mapped_column(String(255))
    department:      Mapped[str]      = mapped_column(String(100),  nullable=True)
    avatar_color:    Mapped[str]      = mapped_column(String(7),    default="#4f46e5")
    avatar_url:      Mapped[str]      = mapped_column(String(500),  nullable=True)
    is_active:       Mapped[int]      = mapped_column(SmallInteger, default=1)
    is_admin:        Mapped[int]      = mapped_column(SmallInteger, default=0)
    organization_id: Mapped[int]      = mapped_column(Integer,      nullable=True)
    phone:           Mapped[str]      = mapped_column(String(20),   nullable=True)
    storage_used_mb: Mapped[float]    = mapped_column(Float,        default=0)
    created_at:      Mapped[datetime] = mapped_column(DateTime,     server_default=func.now())
    last_login:      Mapped[datetime] = mapped_column(DateTime,     nullable=True)