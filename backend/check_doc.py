import sys
sys.path.insert(0, '.')
from dotenv import load_dotenv
load_dotenv()
import asyncio
from app.database import SessionLocal
from app.models.document import Document
from app.lib.document_storage import DocumentStorage

async def check():
    async with SessionLocal() as db:
        doc = await db.get(Document, '1cd6c985-7b4b-40cd-a088-b4abc1ce8848')
        if not doc:
            print('Document not found in DB')
            return
        
        print('Title:', doc.title)
        print('File path:', doc.file_path)
        print('File size:', doc.file_size)
        print('Version:', doc.content_version)
        print('Doc type:', doc.doc_type)
        print('Owner:', doc.owner_id)
        
        try:
            content = await asyncio.to_thread(
                DocumentStorage.load,
                doc.owner_id, doc.id,
                doc_type=doc.doc_type,
                file_path=doc.file_path
            )
            import json
            p = json.loads(content)
            sheets = p.get('_importedSheets', [])
            print('Has _importedSheets:', '_importedSheets' in p)
            print('Sheet count:', len(sheets))
            print('Sheet names:', [s.get('name','?') for s in sheets])
        except Exception as e:
            print('Error loading from R2:', e)

asyncio.run(check())
