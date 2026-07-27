import os
import json
import logging
import traceback
import uuid

from dotenv import load_dotenv
from contextlib import asynccontextmanager

from fastapi import FastAPI, WebSocket, WebSocketDisconnect, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from app.database import engine, Base
from app.api import documents, export, auth, chat

load_dotenv(override=True)

logger = logging.getLogger("uvicorn.error")


def _build_allowed_origins() -> list[str]:
    origins = {"http://localhost:4200", "http://127.0.0.1:4200"}
    env_url = os.getenv("FRONTEND_URL", "").strip()
    if env_url:
        origins.add(env_url)
    # Allow any LAN IP on port 4200 via wildcard not supported by CORSMiddleware,
    # so we also allow the explicit LAN host if set.
    extra = os.getenv("EXTRA_ORIGINS", "")  # comma-separated, e.g. http://192.168.0.10:4200
    for o in extra.split(","):
        o = o.strip()
        if o:
            origins.add(o)
    return list(origins)



import time
import asyncio
from app.lib.document_storage import DocumentStorage
from app.database import SessionLocal
from app.models.document import Document

class ConnectionManager:
    def __init__(self):
        self.rooms: dict[str, dict[str, WebSocket]] = {}
        self.doc_states: dict[str, dict] = {}

    async def _ensure_state(self, doc_id: str):
        if doc_id not in self.doc_states:
            async with SessionLocal() as db:
                doc = await db.get(Document, doc_id)
                content_str = "{}"
                if doc and doc.file_path:
                    try:
                        content_str = await asyncio.to_thread(
                            DocumentStorage.load,
                            doc.owner_id, doc.id,
                            doc.doc_type, doc.file_path
                        )
                    except Exception:
                        pass
                
                state_data = []
                try:
                    state_data = json.loads(content_str)
                except Exception:
                    pass
                if isinstance(state_data, dict):
                    state_data = [state_data]
                    
                self.doc_states[doc_id] = {
                    "seq": 0,
                    "data": state_data,
                    "dirty": False,
                    "last_save": time.time(),
                    "doc_info": {
                        "id": doc.id if doc else doc_id,
                        "owner_id": doc.owner_id if doc else 0,
                        "doc_type": doc.doc_type if doc else "sheet"
                    }
                }

    async def connect(self, doc_id: str, ws: WebSocket) -> str:
        await ws.accept()
        await self._ensure_state(doc_id)
        
        client_id = str(uuid.uuid4())
        self.rooms.setdefault(doc_id, {})[client_id] = ws
        
        # Send initial full state
        state = self.doc_states[doc_id]
        
        content_str = json.dumps(state["data"][0]) if isinstance(state["data"], list) and state["data"] else json.dumps(state["data"])

        payload = json.dumps({
            "type": "update",
            "content": content_str,
            "seq": state["seq"]
        })
        await ws.send_text(payload)
        
        return client_id

    def disconnect(self, doc_id: str, client_id: str):
        room = self.rooms.get(doc_id)
        if room and client_id in room:
            del room[client_id]
            if not room:
                del self.rooms[doc_id]
                # IMPORTANT: Do NOT pop doc_states here immediately.
                # Schedule a final save task before clearing it to prevent data loss.
                asyncio.create_task(self._save_and_evict(doc_id))

    async def _save_and_evict(self, doc_id: str):
        """Save dirty state to R2 immediately on last-user-disconnect, then evict from RAM."""
        state = self.doc_states.get(doc_id)
        if not state:
            return
        # Always save on eviction if there's any data, even if not dirty,
        # to avoid any edge case where the flag was cleared but data wasn't saved.
        doc_info = state.get("doc_info", {})
        owner_id = doc_info.get("owner_id", 0)
        doc_id_str = doc_info.get("id", doc_id)
        doc_type = doc_info.get("doc_type", "sheet")
        if owner_id and doc_id_str and state.get("data"):
            try:
                content_str = json.dumps(state["data"][0]) if isinstance(state["data"], list) and state["data"] else json.dumps(state["data"])

                # SAFETY: Don't save empty sheet data over real R2 content.
                if doc_type == "sheet" and not state.get("dirty"):
                    # If state was never dirtied, the loaded R2 data hasn't changed — skip save.
                    logger.info(f"[EVICT] State not dirty for {doc_id}, skipping R2 write.")
                else:
                    result = await asyncio.to_thread(
                        DocumentStorage.save,
                        owner_id, doc_id_str,
                        content_str, doc_type=doc_type
                    )
                    # Also update the DB file_path and file_size so the next load always
                    # resolves to the correct R2 key, even if the HTTP save was never called.
                    try:
                        async with SessionLocal() as db:
                            doc = await db.get(Document, doc_id_str)
                            if doc:
                                doc.file_path = result["relative_path"]
                                doc.file_size = result["size"]
                                doc.content_version = (doc.content_version or 1) + 1
                                await db.commit()
                    except Exception as db_err:
                        logger.warning(f"[EVICT] DB update failed for {doc_id}: {db_err}")
                    logger.info(f"[EVICT] Final save on disconnect for doc {doc_id} ({result['size']} bytes)")
            except Exception as e:
                logger.error(f"[EVICT] Failed to save doc {doc_id} on disconnect: {e}")
        # Now it's safe to evict from RAM
        self.doc_states.pop(doc_id, None)

    async def broadcast(self, doc_id: str, message: str, sender_id: str | None = None):
        room = self.rooms.get(doc_id)
        if not room:
            return
        dead = []
        for cid, ws in list(room.items()):
            if cid == sender_id:
                continue
            try:
                await ws.send_text(message)
            except Exception:
                dead.append(cid)
        for cid in dead:
            self.disconnect(doc_id, cid)

    def user_count(self, doc_id: str) -> int:
        return len(self.rooms.get(doc_id, {}))


    async def save_if_dirty(self, doc_id: str):
        state = self.doc_states.get(doc_id)
        if not state or not state["dirty"]:
            return
        now = time.time()
        if now - state["last_save"] < 5.0:
            return
            
        state["dirty"] = False
        state["last_save"] = now
        
        doc_info = state["doc_info"]
        content_str = json.dumps(state["data"][0]) if isinstance(state["data"], list) and state["data"] else json.dumps(state["data"])
        try:
            await asyncio.to_thread(
                DocumentStorage.save,
                doc_info["owner_id"], doc_info["id"],
                content_str, doc_type=doc_info["doc_type"]
            )
        except Exception:
            pass

    async def _autosave_loop(self):
        while True:
            await asyncio.sleep(5)
            doc_ids = list(self.doc_states.keys())
            for doc_id in doc_ids:
                try:
                    await self.save_if_dirty(doc_id)
                except Exception:
                    pass

manager = ConnectionManager()




@asynccontextmanager
async def lifespan(app: FastAPI):
    import asyncio
    asyncio.create_task(manager._autosave_loop())
    async with engine.begin() as conn:
        await conn.run_sync(Base.metadata.create_all)

    # Pre-warm the connection pool so first requests don't pay ~300ms cold-connect cost
    async def _warm():
        from sqlalchemy import text
        async with engine.connect() as c:
            await c.execute(text("SELECT 1"))
    try:
        await asyncio.gather(*[_warm() for _ in range(3)])
        logger.info("Connection pool warmed up (3 connections)")
    except Exception as e:
        logger.warning(f"Pool warmup failed (non-fatal): {e}")

    yield


app = FastAPI(title="Office Suite API", lifespan=lifespan)

# CORS must be added BEFORE the error-catcher middleware so it runs on
# error responses too.
app.add_middleware(
    CORSMiddleware,
    allow_origins=_build_allowed_origins(),
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.middleware("http")
async def critical_error_catcher(request: Request, call_next):
    try:
        return await call_next(request)
    except Exception as e:
        logger.exception(f"Unhandled error: {request.method} {request.url}")
        return JSONResponse(
            status_code=500,
            content={"detail": f"Internal server error: {type(e).__name__}: {e}"},
        )


app.include_router(auth.router,      prefix="/api/auth",      tags=["auth"])
app.include_router(documents.router, prefix="/api/documents", tags=["documents"])
app.include_router(export.router,    prefix="/api",           tags=["export"])
app.include_router(chat.router,      prefix="/api/chat",      tags=["chat"])


@app.websocket("/ws/{doc_id}")
async def websocket_endpoint(websocket: WebSocket, doc_id: str):
    try:
        client_id = await manager.connect(doc_id, websocket)
        count = manager.user_count(doc_id)
        presence = json.dumps({"type": "presence", "users": count})
        await manager.broadcast(doc_id, presence, sender_id=client_id)
        await websocket.send_text(presence)
    except WebSocketDisconnect:
        return
    except Exception as e:
        logger.error(f"WebSocket connect error for {doc_id}: {e}")
        return

    try:
        while True:
            data = await websocket.receive_text()
            msg = json.loads(data)
            msg_type = msg.get("type")


            if msg_type == "update":
                content_payload = msg.get("content")
                
                # Keep state updated for new joiners
                state = manager.doc_states.get(doc_id)
                if state:
                    try:
                        parsed = json.loads(content_payload)
                        # SAFETY: Don't overwrite existing state with empty sheet data.
                        # This prevents the race condition where the frontend sends empty
                        # cells before it has loaded the real data from the HTTP API.
                        existing_data = state.get("data", [])
                        if existing_data and isinstance(existing_data, list) and len(existing_data) > 0:
                            existing_first = existing_data[0] if isinstance(existing_data[0], dict) else {}
                            existing_sheets = existing_first.get("_importedSheets", [])
                            existing_has_cells = False
                            for es in existing_sheets:
                                sc = es.get("cells", {})
                                if isinstance(sc, dict) and sc:
                                    existing_has_cells = True
                                    break
                            if existing_has_cells:
                                # Check if incoming data has cells
                                new_sheets = parsed.get("_importedSheets", [])
                                new_has_cells = False
                                for ns in new_sheets:
                                    nsc = ns.get("cells", {})
                                    if isinstance(nsc, dict) and nsc:
                                        new_has_cells = True
                                        break
                                if not new_has_cells:
                                    logger.warning(f"[WS GUARD] Blocked empty sheet update for {doc_id} (existing has data)")
                                    # Still broadcast so other clients see the attempted change,
                                    # but do NOT update doc_states or mark dirty
                                    payload = json.dumps({
                                        "type":    "update",
                                        "content": content_payload,
                                        "title":   msg.get("title"),
                                        "users":   manager.user_count(doc_id),
                                    })
                                    await manager.broadcast(doc_id, payload, sender_id=client_id)
                                    continue

                        state["data"] = [parsed]
                        if msg.get("autosave", True):
                            state["dirty"] = True
                    except Exception:
                        pass

                payload = json.dumps({
                    "type":    "update",
                    "content": content_payload,
                    "title":   msg.get("title"),
                    "users":   manager.user_count(doc_id),
                })
                await manager.broadcast(doc_id, payload, sender_id=client_id)
                
            elif msg_type == "cell_update":
                sheet_idx = msg.get("sheetIdx", 0)
                r = msg.get("r")
                c = msg.get("c")
                value = msg.get("value")
                formatting = msg.get("formatting")
                
                state = manager.doc_states.get(doc_id)
                if state:
                    # Root doc is state["data"][0] if it's a list
                    root_doc = state["data"][0] if isinstance(state["data"], list) and state["data"] else state.get("data", {})
                    
                    if "_importedSheets" not in root_doc:
                        root_doc["_importedSheets"] = []
                    
                    sheets_arr = root_doc["_importedSheets"]
                    while len(sheets_arr) <= sheet_idx:
                        sheets_arr.append({})
                        
                    sheet = sheets_arr[sheet_idx]
                    
                    if "cells" not in sheet:
                        sheet["cells"] = {}
                    if str(r) not in sheet["cells"]:
                        sheet["cells"][str(r)] = {}
                    sheet["cells"][str(r)][str(c)] = value
                    
                    if "formats" not in sheet:
                        sheet["formats"] = {}
                    fmt_key = f"{r},{c}"
                    if formatting:
                        sheet["formats"][fmt_key] = formatting
                    else:
                        sheet["formats"].pop(fmt_key, None)
                        
                    state["seq"] += 1
                    state["dirty"] = True
                    
                    payload = json.dumps({
                        "type": "cell_update",
                        "sheetIdx": sheet_idx,
                        "r": r,
                        "c": c,
                        "value": value,
                        "formatting": formatting,
                        "seq": state["seq"]
                    })
                    await manager.broadcast(doc_id, payload, sender_id=client_id)


            elif msg_type == "cursor":
                payload = json.dumps({
                    "type":      "cursor",
                    "client_id": client_id,
                    "r":         msg.get("r"),
                    "c":         msg.get("c"),
                })
                await manager.broadcast(doc_id, payload, sender_id=client_id)

    except WebSocketDisconnect:
        manager.disconnect(doc_id, client_id)
        await manager.broadcast(
            doc_id,
            json.dumps({"type": "cursor_remove", "client_id": client_id}),
        )
        await manager.broadcast(
            doc_id,
            json.dumps({"type": "presence", "users": manager.user_count(doc_id)}),
        )