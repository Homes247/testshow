from sqlalchemy import Integer, String, ForeignKey
from sqlalchemy.orm import Mapped, mapped_column
from app.database import Base


class UserCredential(Base):
    """
    Maps to VMail's user_credentials table.
    login_email is the actual login identifier — users.email may differ.
    Do NOT let SQLAlchemy create or modify this table.
    """
    __tablename__ = "user_credentials"
    __table_args__ = {"extend_existing": True}

    id:          Mapped[int] = mapped_column(Integer, primary_key=True)
    user_id:     Mapped[int] = mapped_column(Integer, ForeignKey("users.id", ondelete="CASCADE"), nullable=False)
    login_email: Mapped[str] = mapped_column(String(100), unique=True, nullable=False, index=True)