import asyncio
import os
import sys

# Add backend dir to python path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from sqlalchemy.ext.asyncio import create_async_engine
from sqlalchemy import text
from app.database import DATABASE_URL
from app.lib.document_storage import DocumentStorage

async def main():
    engine = create_async_engine(DATABASE_URL)
    
    async with engine.begin() as conn:
        print("Checking if columns exist...")
        # Check if columns exist
        result = await conn.execute(text("SHOW COLUMNS FROM documents LIKE 'file_path'"))
        if not result.scalar():
            print("Adding file_path, file_size, content_version columns...")
            await conn.execute(text("ALTER TABLE documents ADD COLUMN file_path VARCHAR(500) NULL;"))
            await conn.execute(text("ALTER TABLE documents ADD COLUMN file_size INT NULL;"))
            await conn.execute(text("ALTER TABLE documents ADD COLUMN content_version INT NOT NULL DEFAULT 1;"))
        else:
            print("Columns already exist.")
            
        print("Migrating content to files...")
        result = await conn.execute(text("SELECT id, owner_id, content FROM documents WHERE content IS NOT NULL AND file_path IS NULL"))
        rows = result.mappings().all()
        
        migrated = 0
        failed = 0
        
        for row in rows:
            if not row['content']:
                continue
                
            try:
                storage_res = DocumentStorage.save(row['owner_id'], row['id'], row['content'])
                await conn.execute(
                    text("UPDATE documents SET file_path = :file_path, file_size = :file_size WHERE id = :id"),
                    {"file_path": storage_res["relative_path"], "file_size": storage_res["size"], "id": row['id']}
                )
                migrated += 1
                print(f"[OK] {row['id']}")
            except Exception as e:
                failed += 1
                print(f"[FAIL] {row['id']}: {e}")
                
        print(f"\nDone. Migrated: {migrated}, Failed: {failed}")
        
    await engine.dispose()

if __name__ == "__main__":
    asyncio.run(main())
