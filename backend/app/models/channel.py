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


class Channel(Base):
    __tablename__ = "channels"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String(100), nullable=False)
    description = Column(String(500), nullable=True)
    is_private = Column(SmallInteger, default=0)
    created_by = Column(Integer, ForeignKey("users.id", ondelete="CASCADE"), nullable=False)
    organization_id = Column(Integer, ForeignKey("organizations.id", ondelete="CASCADE"), nullable=True, index=True)
    avatar_url = Column(String(500), nullable=True)
    created_at = Column(DateTime, default=now_ist)
    creator = relationship("User", foreign_keys=[created_by])
    members = relationship("ChannelMember", back_populates="channel", cascade="all, delete")
    messages = relationship("ChannelMessage", back_populates="channel", cascade="all, delete")


class ChannelMember(Base):
    __tablename__ = "channel_members"
    id = Column(Integer, primary_key=True, index=True)
    channel_id = Column(Integer, ForeignKey("channels.id", ondelete="CASCADE"), nullable=False)
    user_id = Column(Integer, ForeignKey("users.id", ondelete="CASCADE"), nullable=False)
    is_admin = Column(SmallInteger, default=0)
    joined_at = Column(DateTime, default=now_ist)
    channel = relationship("Channel", back_populates="members")
    user = relationship("User", foreign_keys=[user_id])


class ChannelMessage(Base):
    __tablename__ = "channel_messages"
    id = Column(Integer, primary_key=True, index=True)
    channel_id = Column(Integer, ForeignKey("channels.id", ondelete="CASCADE"), nullable=False)
    sender_id = Column(Integer, ForeignKey("users.id", ondelete="CASCADE"), nullable=False)
    message = Column(Text, nullable=False)
    attachment_name = Column(String(255), nullable=True)
    attachment_size = Column(Integer, nullable=True)
    is_file = Column(SmallInteger, default=0)
    file_path = Column(String(500), nullable=True)
    is_edited = Column(SmallInteger, default=0)
    delete_status = Column(SmallInteger, default=0)
    reactions = Column(Text, nullable=True, default="{}")
    created_at = Column(DateTime, default=now_ist)
    updated_at = Column(DateTime, default=now_ist, onupdate=now_ist)
    channel = relationship("Channel", back_populates="messages")
    sender = relationship("User", foreign_keys=[sender_id])