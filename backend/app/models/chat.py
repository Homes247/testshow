from sqlalchemy import Column, Integer, Text, DateTime, SmallInteger, ForeignKey, String
from sqlalchemy.orm import relationship
from app.database import Base
from datetime import datetime
from zoneinfo import ZoneInfo  # built-in from Python 3.9+

# ─────────────────────────────────────────────
# IST time using system zoneinfo database
# Zero network calls — instant, always correct
# ZoneInfo("Asia/Kolkata") reads the timezone
# rules from the OS, not the server clock,
# so deployment location doesn't matter at all.
# ─────────────────────────────────────────────

IST = ZoneInfo("Asia/Kolkata")

def now_ist() -> datetime:
    return datetime.now(IST).replace(tzinfo=None)  # naive IST for MySQL DateTime


class ChatConversation(Base):
    __tablename__ = "chat_conversations"
    id = Column(Integer, primary_key=True, index=True)
    user1_id = Column(Integer, ForeignKey("users.id", ondelete="CASCADE"), nullable=False)
    user2_id = Column(Integer, ForeignKey("users.id", ondelete="CASCADE"), nullable=False)
    created_at = Column(DateTime, default=now_ist)
    updated_at = Column(DateTime, default=now_ist, onupdate=now_ist)
    user1 = relationship("User", foreign_keys=[user1_id])
    user2 = relationship("User", foreign_keys=[user2_id])
    messages = relationship("ChatMessage", back_populates="conversation", cascade="all, delete")


class ChatMessage(Base):
    __tablename__ = "chat_messages"
    id = Column(Integer, primary_key=True, index=True)
    conversation_id = Column(Integer, ForeignKey("chat_conversations.id", ondelete="CASCADE"), nullable=False)
    sender_id = Column(Integer, ForeignKey("users.id", ondelete="CASCADE"), nullable=False)
    message = Column(Text, nullable=False)
    is_file = Column(SmallInteger, default=0)
    file_path = Column(String(500), nullable=True)
    is_read = Column(SmallInteger, default=0)
    is_edited = Column(SmallInteger, default=0)
    delete_status = Column(SmallInteger, default=0)
    reactions = Column(Text, nullable=True, default="{}")
    created_at = Column(DateTime, default=now_ist)
    updated_at = Column(DateTime, default=now_ist, onupdate=now_ist)
    conversation = relationship("ChatConversation", back_populates="messages")
    sender = relationship("User", foreign_keys=[sender_id])