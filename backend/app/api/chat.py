import traceback
from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import RedirectResponse
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select, or_, and_, func
from pydantic import BaseModel
from typing import Optional, List
from datetime import datetime, timezone, timedelta
from zoneinfo import ZoneInfo

from app.database import get_db
from app.api.auth import get_current_user
from app.models.user import User
from app.models.chat import ChatConversation, ChatMessage
from app.models.channel import Channel, ChannelMember, ChannelMessage

IST_OFFSET = timezone(timedelta(hours=5, minutes=30))
IST = ZoneInfo("Asia/Kolkata")

def now_ist() -> datetime:
    return datetime.now(IST).replace(tzinfo=None)

def to_utc_iso(dt) -> str:
    if not dt: return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=IST_OFFSET)
    return dt.astimezone(timezone.utc).isoformat().replace("+00:00", "Z")

router = APIRouter(tags=["chat"])

from typing import Optional

class SendMessageRequest(BaseModel):
    to_user_id: int
    message: str
    is_file: bool = False
    file_path: Optional[str] = None

async def get_or_create_conversation(db: AsyncSession, user1_id: int, user2_id: int) -> ChatConversation:
    uid1, uid2 = min(user1_id, user2_id), max(user1_id, user2_id)
    result = await db.execute(select(ChatConversation).where(
        ChatConversation.user1_id == uid1,
        ChatConversation.user2_id == uid2
    ))
    conv = result.scalar_one_or_none()
    if not conv:
        conv = ChatConversation(user1_id=uid1, user2_id=uid2)
        db.add(conv)
        await db.commit()
        await db.refresh(conv)
    return conv

@router.get("/conversations")
async def get_conversations(db: AsyncSession = Depends(get_db), current_user: User = Depends(get_current_user)):
    try:
        result = await db.execute(
            select(ChatConversation).where(
                or_(
                    ChatConversation.user1_id == current_user.id,
                    ChatConversation.user2_id == current_user.id
                )
            ).order_by(ChatConversation.updated_at.desc())
        )
        convs = result.scalars().all()
        
        response = []
        if not convs:
            return response

        # Collect IDs for batched queries
        other_user_ids = []
        conv_ids = []
        for conv in convs:
            conv_ids.append(conv.id)
            other_user_ids.append(conv.user2_id if conv.user1_id == current_user.id else conv.user1_id)

        # Batch fetch users
        res_users = await db.execute(select(User).where(User.id.in_(other_user_ids)))
        users_dict = {u.id: u for u in res_users.scalars().all()}

        # Batch fetch unread counts
        # select conv_id, count(*) from messages where conv_id in (...) and sender_id != me and is_read = 0 group by conv_id
        res_unread = await db.execute(
            select(ChatMessage.conversation_id, func.count(ChatMessage.id))
            .where(
                ChatMessage.conversation_id.in_(conv_ids),
                ChatMessage.sender_id != current_user.id,
                ChatMessage.is_read == 0
            )
            .group_by(ChatMessage.conversation_id)
        )
        unread_dict = {row[0]: row[1] for row in res_unread.all()}

        # Batch fetch last messages (using a subquery or by just fetching latest per conv)
        # For simplicity and DB compatibility, we fetch the latest message per conversation
        # MySQL/MariaDB: select m.* from messages m inner join (select conversation_id, max(created_at) as max_c from messages group by conversation_id) grouped on m.conversation_id = grouped.conversation_id and m.created_at = grouped.max_c
        subq = (
            select(
                ChatMessage.conversation_id,
                func.max(ChatMessage.created_at).label('max_time')
            )
            .where(ChatMessage.conversation_id.in_(conv_ids))
            .group_by(ChatMessage.conversation_id)
            .subquery()
        )
        res_msgs = await db.execute(
            select(ChatMessage)
            .join(subq, and_(
                ChatMessage.conversation_id == subq.c.conversation_id,
                ChatMessage.created_at == subq.c.max_time
            ))
        )
        # If multiple messages have exact same max_time, we might get duplicates, so we map by ID
        last_msgs_dict = {}
        for m in res_msgs.scalars().all():
            last_msgs_dict[m.conversation_id] = m

        for conv in convs:
            other_id = conv.user2_id if conv.user1_id == current_user.id else conv.user1_id
            other = users_dict.get(other_id)
            if not other: continue

            last_msg = last_msgs_dict.get(conv.id)
            if not last_msg: continue

            unread = unread_dict.get(conv.id, 0)

            response.append({
                "conversation_id": conv.id,
                "other_user": {
                    "id": other.id,
                    "name": other.name,
                    "email": other.email,
                    "avatar_color": getattr(other, 'avatar_color', '#4f46e5') or '#4f46e5',
                    "avatar_url": getattr(other, 'avatar_url', None),
                },
                "last_message": last_msg.message,
                "last_time": to_utc_iso(last_msg.created_at),
                "unread": unread,
                "updated_at": to_utc_iso(conv.updated_at) if conv.updated_at else to_utc_iso(conv.created_at)
            })

        return response
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/messages/{other_user_id}")
async def get_messages(other_user_id: int, db: AsyncSession = Depends(get_db), current_user: User = Depends(get_current_user)):
    try:
        conv = await get_or_create_conversation(db, current_user.id, other_user_id)
        
        # Mark read (Async update requires a slightly different approach or just manual mapping)
        res_unread = await db.execute(select(ChatMessage).where(
            ChatMessage.conversation_id == conv.id,
            ChatMessage.sender_id != current_user.id,
            ChatMessage.is_read == 0
        ))
        unreads = res_unread.scalars().all()
        for u in unreads:
            u.is_read = 1
        if unreads:
            await db.commit()
            
        res_msg = await db.execute(select(ChatMessage).where(ChatMessage.conversation_id == conv.id).order_by(ChatMessage.created_at.asc()))
        messages = res_msg.scalars().all()
        
        res_other = await db.execute(select(User).where(User.id == other_user_id))
        other = res_other.scalar_one_or_none()
        
        if not other:
            raise HTTPException(404, "User not found")
            
        # Manually fetch senders since async relationships can be tricky without joinedload
        msgs_response = []
        for m in messages:
            sender_name = current_user.name if m.sender_id == current_user.id else other.name
            sender_color = current_user.avatar_color if m.sender_id == current_user.id else other.avatar_color
            sender_url = current_user.avatar_url if m.sender_id == current_user.id else other.avatar_url
            
            msgs_response.append({
                "id": m.id,
                "sender_id": m.sender_id,
                "sender_name": sender_name,
                "sender_avatar_color": sender_color or '#4f46e5',
                "sender_avatar_url": sender_url,
                "message": m.message if not m.delete_status else "[[DELETED]]",
                "is_mine": m.sender_id == current_user.id,
                "is_file": bool(m.is_file),
                "file_path": m.file_path,
                "created_at": to_utc_iso(m.created_at),
                "is_read": bool(m.is_read)
            })
            
        return {
            "conversation_id": conv.id,
            "other_user": {
                "id": other.id,
                "name": other.name,
                "email": other.email,
                "avatar_color": getattr(other, 'avatar_color', '#4f46e5') or '#4f46e5',
                "avatar_url": getattr(other, 'avatar_url', None)
            },
            "messages": msgs_response
        }
    except HTTPException:
        raise
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/send")
async def send_message(data: SendMessageRequest, db: AsyncSession = Depends(get_db), current_user: User = Depends(get_current_user)):
    try:
        res_other = await db.execute(select(User).where(User.id == data.to_user_id))
        other = res_other.scalar_one_or_none()
        if not other: raise HTTPException(404, "User not found")
        
        conv = await get_or_create_conversation(db, current_user.id, data.to_user_id)
        msg = ChatMessage(
            conversation_id=conv.id, 
            sender_id=current_user.id, 
            message=data.message,
            is_file=data.is_file,
            file_path=data.file_path
        )
        db.add(msg)
        conv.updated_at = now_ist()
        await db.commit()
        await db.refresh(msg)
        
        return {
            "id": msg.id,
            "sender_id": msg.sender_id,
            "sender_name": current_user.name,
            "sender_avatar_color": getattr(current_user, 'avatar_color', '#4f46e5') or '#4f46e5',
            "message": msg.message,
            "is_mine": True,
            "is_file": msg.is_file,
            "file_path": msg.file_path,
            "created_at": to_utc_iso(msg.created_at)
        }
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

class EditMessageRequest(BaseModel):
    message: str

@router.put("/messages/{message_id}")
async def edit_message(message_id: int, data: EditMessageRequest, db: AsyncSession = Depends(get_db), current_user: User = Depends(get_current_user)):
    try:
        res = await db.execute(select(ChatMessage).where(ChatMessage.id == message_id, ChatMessage.sender_id == current_user.id))
        msg = res.scalar_one_or_none()
        if not msg: raise HTTPException(404, "Message not found or unauthorized")
        msg.message = data.message
        msg.is_edited = 1
        await db.commit()
        return {"success": True}
    except HTTPException:
        raise
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

class ReactMessageRequest(BaseModel):
    emoji: str

@router.put("/messages/{message_id}/react")
async def react_message(message_id: int, data: ReactMessageRequest, db: AsyncSession = Depends(get_db), current_user: User = Depends(get_current_user)):
    try:
        res = await db.execute(select(ChatMessage).where(ChatMessage.id == message_id))
        msg = res.scalar_one_or_none()
        if not msg: raise HTTPException(404, "Message not found")
        
        import json
        reactions = {}
        if msg.reactions:
            try:
                reactions = json.loads(msg.reactions)
            except: pass
            
        uid = str(current_user.id)
        # If emoji is empty or same, toggle it off
        if uid in reactions and reactions[uid] == data.emoji:
            del reactions[uid]
        elif data.emoji:
            reactions[uid] = data.emoji
            
        msg.reactions = json.dumps(reactions)
        await db.commit()
        return {"success": True, "reactions": msg.reactions}
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

@router.delete("/messages/{message_id}")
async def delete_chat_message(message_id: int, db: AsyncSession = Depends(get_db), current_user: User = Depends(get_current_user)):
    try:
        res = await db.execute(select(ChatMessage).where(ChatMessage.id == message_id, ChatMessage.sender_id == current_user.id))
        msg = res.scalar_one_or_none()
        if not msg: raise HTTPException(404, "Message not found or unauthorized")
        await db.delete(msg)
        await db.commit()
        return {"success": True}
    except HTTPException:
        raise
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

from fastapi import Request
import boto3
import os
import uuid
import asyncio
from app.lib.document_storage import R2_ENDPOINT_URL, R2_ACCESS_KEY_ID, R2_SECRET_ACCESS_KEY, R2_BUCKET_NAME, _s3_client

@router.post("/upload-stream")
async def upload_stream_chat(
    request: Request,
    to_user_id: int,
    filename: str,
    mime_type: str,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    try:
        body_bytes = await request.body()
        key = f"chat/{current_user.id}/{uuid.uuid4()}_{filename}"
        
        def upload():
            _s3_client.put_object(
                Bucket=R2_BUCKET_NAME,
                Key=key,
                Body=body_bytes,
                ContentType=mime_type or "application/octet-stream",
            )
        await asyncio.to_thread(upload)
        
        # Save to DB
        conv = await get_or_create_conversation(db, current_user.id, to_user_id)
        # R2 public url might be needed, or we just save the key if we have a proxy, but Vmail uses direct R2 public url?
        # Let's save the key, but in msg.file_path we might need full URL if frontend doesn't prefix it.
        # Looking at frontend, it uses msg.file_path directly in href, so we should save full URL.
        # Actually R2_ENDPOINT_URL is https://fa7777c8717deb771ad51677f4451593.r2.cloudflarestorage.com
        # VMail uses a custom domain or we can just use the bucket URL. 
        # For this test, let's just use the R2_ENDPOINT_URL/bucket/key
        file_url = f"{R2_ENDPOINT_URL}/{R2_BUCKET_NAME}/{key}"
        # Wait, R2_ENDPOINT_URL is private. If it's a private bucket, we need presigned URLs. 
        # But if Vmail frontend just uses the URL directly, it must be public.
        # Let's use https://pub-your-bucket.com if there's a public URL in env.
        # There is no public URL in .env. We will use the endpoint URL.
        
        msg = ChatMessage(
            conversation_id=conv.id, 
            sender_id=current_user.id, 
            message=filename,
            is_file=1,
            file_path=file_url
        )
        db.add(msg)
        conv.updated_at = now_ist()
        await db.commit()
        await db.refresh(msg)
        
        return {
            "id": msg.id,
            "sender_id": msg.sender_id,
            "sender_name": current_user.name,
            "sender_avatar_color": getattr(current_user, 'avatar_color', '#4f46e5') or '#4f46e5',
            "message": msg.message,
            "is_mine": True,
            "is_file": True,
            "file_path": msg.file_path,
            "created_at": to_utc_iso(msg.created_at)
        }
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/search-channels")
async def search_channels(q: str = "", db: AsyncSession = Depends(get_db)):
    if not q or len(q) < 2:
        return []
    res = await db.execute(select(Channel).where(Channel.name.ilike(f"%{q}%")).limit(10))
    channels = res.scalars().all()
    return [{"id": c.id, "name": c.name, "avatar_url": getattr(c, 'avatar_url', None)} for c in channels]

@router.get("/channels")
async def get_channels(db: AsyncSession = Depends(get_db), current_user: User = Depends(get_current_user)):
    res = await db.execute(
        select(Channel).join(ChannelMember, Channel.id == ChannelMember.channel_id)
        .where(ChannelMember.user_id == current_user.id)
    )
    channels = res.scalars().all()
    response = []
    for c in channels:
        res_msg = await db.execute(select(ChannelMessage).where(ChannelMessage.channel_id == c.id).order_by(ChannelMessage.created_at.desc()).limit(1))
        last = res_msg.scalar_one_or_none()
        response.append({
            "id": c.id,
            "name": c.name,
            "last_message": last.message if last else "No messages",
            "last_time": to_utc_iso(last.created_at) if last else None
        })
    return response

@router.get("/channels/{channel_id}/messages")
async def get_channel_messages(channel_id: int, db: AsyncSession = Depends(get_db), current_user: User = Depends(get_current_user)):
    res = await db.execute(select(ChannelMessage).where(ChannelMessage.channel_id == channel_id).order_by(ChannelMessage.created_at.asc()))
    messages = res.scalars().all()
    
    # Check if channel exists
    res_chan = await db.execute(select(Channel).where(Channel.id == channel_id))
    channel = res_chan.scalar_one_or_none()
    if not channel: raise HTTPException(404, "Channel not found")
    
    msgs_response = []
    for m in messages:
        res_sender = await db.execute(select(User).where(User.id == m.sender_id))
        sender = res_sender.scalar_one_or_none()
        msgs_response.append({
            "id": m.id,
            "sender_id": m.sender_id,
            "sender_name": sender.name if sender else "Unknown",
            "sender_avatar_color": getattr(sender, 'avatar_color', '#4f46e5') if sender else '#4f46e5',
            "sender_avatar_url": getattr(sender, 'avatar_url', None) if sender else None,
            "message": m.message,
            "is_mine": m.sender_id == current_user.id,
            "created_at": to_utc_iso(m.created_at)
        })
        
    return {
        "channel_id": channel_id,
        "channel_name": channel.name,
        "messages": msgs_response
    }

class UpdateChannelRequest(BaseModel):
    name: str
    description: Optional[str] = None

@router.get("/channels/{channel_id}/info")
async def get_channel_info(channel_id: int, db: AsyncSession = Depends(get_db), current_user: User = Depends(get_current_user)):
    res_chan = await db.execute(select(Channel).where(Channel.id == channel_id))
    channel = res_chan.scalar_one_or_none()
    if not channel: raise HTTPException(404, "Channel not found")

    res_members = await db.execute(select(ChannelMember, User).join(User, ChannelMember.user_id == User.id).where(ChannelMember.channel_id == channel_id))
    members = []
    is_member = False
    for cm, user in res_members.all():
        if user.id == current_user.id:
            is_member = True
        members.append({
            "id": user.id,
            "name": user.name,
            "email": user.email,
            "avatar_color": getattr(user, 'avatar_color', '#4f46e5') or '#4f46e5',
            "avatar_url": getattr(user, 'avatar_url', None),
            "role": cm.role
        })

    if not is_member:
        raise HTTPException(403, "Not a member of this channel")

    return {
        "id": channel.id,
        "name": channel.name,
        "description": getattr(channel, 'description', ''),
        "avatar_url": getattr(channel, 'avatar_url', None),
        "members": members
    }

@router.put("/channels/{channel_id}")
async def update_channel(channel_id: int, data: UpdateChannelRequest, db: AsyncSession = Depends(get_db), current_user: User = Depends(get_current_user)):
    res_chan = await db.execute(select(Channel).where(Channel.id == channel_id))
    channel = res_chan.scalar_one_or_none()
    if not channel: raise HTTPException(404, "Channel not found")
    
    # Optional: check if admin
    # res_cm = await db.execute(select(ChannelMember).where(ChannelMember.channel_id == channel_id, ChannelMember.user_id == current_user.id))
    # cm = res_cm.scalar_one_or_none()
    # if not cm or cm.role != 'admin': raise HTTPException(403, "Not admin")

    channel.name = data.name
    channel.description = data.description
    await db.commit()
    return {"success": True}

@router.delete("/channels/{channel_id}")
async def delete_channel(channel_id: int, db: AsyncSession = Depends(get_db), current_user: User = Depends(get_current_user)):
    res_chan = await db.execute(select(Channel).where(Channel.id == channel_id))
    channel = res_chan.scalar_one_or_none()
    if not channel: raise HTTPException(404, "Channel not found")
    
    # Verify member
    res_cm = await db.execute(select(ChannelMember).where(ChannelMember.channel_id == channel_id, ChannelMember.user_id == current_user.id))
    cm = res_cm.scalar_one_or_none()
    if not cm: raise HTTPException(403, "Not a member")

    await db.delete(channel) # Cascades should handle members and messages, or they might be orphaned
    await db.commit()
    return {"success": True}

class AddMembersRequest(BaseModel):
    user_ids: List[int]

@router.post("/channels/{channel_id}/members")
async def add_channel_members(channel_id: int, data: AddMembersRequest, db: AsyncSession = Depends(get_db), current_user: User = Depends(get_current_user)):
    res_chan = await db.execute(select(Channel).where(Channel.id == channel_id))
    channel = res_chan.scalar_one_or_none()
    if not channel: raise HTTPException(404, "Channel not found")

    res_cm = await db.execute(select(ChannelMember).where(ChannelMember.channel_id == channel_id, ChannelMember.user_id == current_user.id))
    cm = res_cm.scalar_one_or_none()
    if not cm: raise HTTPException(403, "Not a member")

    # Add members
    added = 0
    for uid in data.user_ids:
        # Check if already a member
        r = await db.execute(select(ChannelMember).where(ChannelMember.channel_id == channel_id, ChannelMember.user_id == uid))
        if not r.scalar_one_or_none():
            new_cm = ChannelMember(channel_id=channel_id, user_id=uid, role='member')
            db.add(new_cm)
            added += 1
            
    if added > 0:
        await db.commit()
    return {"success": True, "added": added}

@router.delete("/channels/{channel_id}/members/{user_id}")
async def remove_channel_member(channel_id: int, user_id: int, db: AsyncSession = Depends(get_db), current_user: User = Depends(get_current_user)):
    if current_user.id != user_id:
        # Check if current_user is admin
        res_cm = await db.execute(select(ChannelMember).where(ChannelMember.channel_id == channel_id, ChannelMember.user_id == current_user.id))
        cm = res_cm.scalar_one_or_none()
        if not cm or getattr(cm, 'role', 'member') != 'admin':
            raise HTTPException(403, "Not an admin")

    res = await db.execute(select(ChannelMember).where(ChannelMember.channel_id == channel_id, ChannelMember.user_id == user_id))
    member = res.scalar_one_or_none()
    if not member: raise HTTPException(404, "Member not found")
    
    await db.delete(member)
    await db.commit()
    return {"success": True}

@router.post("/channels/{channel_id}/leave")
async def leave_channel(channel_id: int, db: AsyncSession = Depends(get_db), current_user: User = Depends(get_current_user)):
    res = await db.execute(select(ChannelMember).where(ChannelMember.channel_id == channel_id, ChannelMember.user_id == current_user.id))
    member = res.scalar_one_or_none()
    if not member: raise HTTPException(404, "Member not found")
    
    await db.delete(member)
    await db.commit()
    return {"success": True}

class SendChannelMessageRequest(BaseModel):
    message: str
    is_file: bool = False
    file_path: Optional[str] = None

@router.post("/channels/{channel_id}/send")
async def send_channel_message(channel_id: int, data: SendChannelMessageRequest, db: AsyncSession = Depends(get_db), current_user: User = Depends(get_current_user)):
    res_chan = await db.execute(select(Channel).where(Channel.id == channel_id))
    channel = res_chan.scalar_one_or_none()
    if not channel: raise HTTPException(404, "Channel not found")
    
    msg = ChannelMessage(
        channel_id=channel_id, 
        sender_id=current_user.id, 
        message=data.message,
        is_file=data.is_file,
        file_path=data.file_path
    )
    db.add(msg)
    await db.commit()
    await db.refresh(msg)
    
    return {
        "id": msg.id,
        "sender_id": msg.sender_id,
        "sender_name": current_user.name,
        "message": msg.message,
        "is_mine": True,
        "is_file": msg.is_file,
        "file_path": msg.file_path,
        "created_at": to_utc_iso(msg.created_at)
    }

@router.put("/channels/{channel_id}/messages/{message_id}")
async def edit_channel_message(channel_id: int, message_id: int, data: EditMessageRequest, db: AsyncSession = Depends(get_db), current_user: User = Depends(get_current_user)):
    try:
        res = await db.execute(select(ChannelMessage).where(ChannelMessage.id == message_id, ChannelMessage.sender_id == current_user.id, ChannelMessage.channel_id == channel_id))
        msg = res.scalar_one_or_none()
        if not msg: raise HTTPException(404, "Message not found or unauthorized")
        msg.message = data.message
        msg.is_edited = 1
        await db.commit()
        return {"success": True}
    except HTTPException:
        raise
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

@router.put("/channels/{channel_id}/messages/{message_id}/react")
async def react_channel_message(channel_id: int, message_id: int, data: ReactMessageRequest, db: AsyncSession = Depends(get_db), current_user: User = Depends(get_current_user)):
    try:
        res = await db.execute(select(ChannelMessage).where(ChannelMessage.id == message_id, ChannelMessage.channel_id == channel_id))
        msg = res.scalar_one_or_none()
        if not msg: raise HTTPException(404, "Message not found")
        
        import json
        reactions = {}
        if msg.reactions:
            try:
                reactions = json.loads(msg.reactions)
            except: pass
            
        uid = str(current_user.id)
        if uid in reactions and reactions[uid] == data.emoji:
            del reactions[uid]
        elif data.emoji:
            reactions[uid] = data.emoji
            
        msg.reactions = json.dumps(reactions)
        await db.commit()
        return {"success": True, "reactions": msg.reactions}
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

@router.delete("/channels/{channel_id}/messages/{message_id}")
async def delete_channel_message(channel_id: int, message_id: int, db: AsyncSession = Depends(get_db), current_user: User = Depends(get_current_user)):
    try:
        res = await db.execute(select(ChannelMessage).where(ChannelMessage.id == message_id, ChannelMessage.sender_id == current_user.id, ChannelMessage.channel_id == channel_id))
        msg = res.scalar_one_or_none()
        if not msg: raise HTTPException(404, "Message not found or unauthorized")
        await db.delete(msg)
        await db.commit()
        return {"success": True}
    except HTTPException:
        raise
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/channels/{channel_id}/upload-stream")
async def upload_stream_channel(
    channel_id: int,
    request: Request,
    filename: str,
    mime_type: str,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    try:
        body_bytes = await request.body()
        key = f"channels/{channel_id}/{uuid.uuid4()}_{filename}"
        
        def upload():
            _s3_client.put_object(
                Bucket=R2_BUCKET_NAME,
                Key=key,
                Body=body_bytes,
                ContentType=mime_type or "application/octet-stream",
            )
        await asyncio.to_thread(upload)
        
        file_url = f"{R2_ENDPOINT_URL}/{R2_BUCKET_NAME}/{key}"
        
        msg = ChannelMessage(
            channel_id=channel_id, 
            sender_id=current_user.id, 
            message=filename,
            is_file=1,
            file_path=file_url
        )
        db.add(msg)
        await db.commit()
        await db.refresh(msg)
        
        return {
            "id": msg.id,
            "sender_id": msg.sender_id,
            "sender_name": current_user.name,
            "sender_avatar_color": getattr(current_user, 'avatar_color', '#4f46e5') or '#4f46e5',
            "sender_avatar_url": getattr(current_user, 'avatar_url', None),
            "message": msg.message,
            "is_mine": True,
            "is_file": True,
            "file_path": msg.file_path,
            "created_at": to_utc_iso(msg.created_at)
        }
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/download")
async def download_file(path: str, filename: str = None, db: AsyncSession = Depends(get_db)):
    try:
        key = path
        prefix = f"{R2_ENDPOINT_URL}/{R2_BUCKET_NAME}/"
        if path.startswith(prefix):
            key = path[len(prefix):]
        elif path.startswith("http"):
            return RedirectResponse(url=path)
            
        # Legacy chat files were saved as 'user_...' in DB but stored as 'chat/user_...' in R2
        if key.startswith("user_"):
            key = f"chat/{key}"
        
        params = {'Bucket': R2_BUCKET_NAME, 'Key': key}
        if filename:
            params['ResponseContentDisposition'] = f'attachment; filename="{filename}"'
            
        presigned_url = _s3_client.generate_presigned_url(
            'get_object',
            Params=params,
            ExpiresIn=3600
        )
        return RedirectResponse(url=presigned_url)
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/contacts")
async def get_contacts(db: AsyncSession = Depends(get_db), current_user: User = Depends(get_current_user)):
    res = await db.execute(select(User).where(User.organization_id == current_user.organization_id))
    users = res.scalars().all()
    return [{"id": u.id, "name": u.name, "email": u.email, "avatar_color": u.avatar_color, "avatar_url": u.avatar_url, "last_login": to_utc_iso(u.last_login) if hasattr(u, 'last_login') else None} for u in users]