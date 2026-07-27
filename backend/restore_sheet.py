"""
Restore the sheet data from the local backup file (c49c.json) to R2 storage.
This re-uploads the original data that was overwritten by the empty-save race condition.
"""
import json
import os
import sys

# Load environment
from dotenv import load_dotenv
load_dotenv(override=True)

from app.lib.document_storage import DocumentStorage

DOC_ID = "7a217da3-8cfe-4808-8027-c49c3a139817"
BACKUP_FILE = os.path.join(os.path.dirname(__file__), "c49c.json")

def main():
    # Read the backup
    print(f"Reading backup from: {BACKUP_FILE}")
    with open(BACKUP_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
    
    # Verify it has real content
    cells = data.get("cells", {})
    sheets = data.get("_importedSheets", [])
    print(f"  Root cells: {len(cells)} rows")
    print(f"  Imported sheets: {len(sheets)}")
    for i, s in enumerate(sheets):
        sc = s.get("cells", {})
        print(f"    Sheet {i} '{s.get('name','?')}': {len(sc)} rows with data")
    
    if not cells and not sheets:
        print("ERROR: Backup file appears empty, aborting!")
        sys.exit(1)
    
    content_str = json.dumps(data)
    print(f"\nTotal content size: {len(content_str):,} bytes")
    
    # We need the owner_id. Let's get it from the DB.
    import asyncio
    from app.database import SessionLocal
    from app.models.document import Document
    
    async def restore():
        async with SessionLocal() as db:
            doc = await db.get(Document, DOC_ID)
            if not doc:
                print(f"ERROR: Document {DOC_ID} not found in database!")
                sys.exit(1)
            
            print(f"\nDocument found:")
            print(f"  Title: {doc.title}")
            print(f"  Owner ID: {doc.owner_id}")
            print(f"  Current file_path: {doc.file_path}")
            print(f"  Current file_size: {doc.file_size}")
            
            # Save to R2
            print(f"\nUploading to R2...")
            result = DocumentStorage.save(
                doc.owner_id, doc.id, content_str, doc_type=doc.doc_type
            )
            print(f"  Saved! Key: {result['relative_path']}, Size: {result['size']:,} bytes")
            
            # Update DB
            doc.file_path = result["relative_path"]
            doc.file_size = result["size"]
            doc.content_version = (doc.content_version or 1) + 1
            await db.commit()
            print(f"  DB updated. New content_version: {doc.content_version}")
            
            # Verify by loading back
            print(f"\nVerifying by loading from R2...")
            loaded = DocumentStorage.load(
                doc.owner_id, doc.id, doc_type=doc.doc_type, file_path=doc.file_path
            )
            loaded_data = json.loads(loaded)
            loaded_sheets = loaded_data.get("_importedSheets", [])
            print(f"  Loaded back {len(loaded)} bytes")
            print(f"  Contains {len(loaded_sheets)} sheets")
            if loaded_sheets:
                print(f"  First sheet name: {loaded_sheets[0].get('name','?')}")
                print(f"  First sheet has {len(loaded_sheets[0].get('cells',{}))} rows with data")
            print("\n✅ RESTORE COMPLETE! Reload the sheet in your browser.")
    
    asyncio.run(restore())

if __name__ == "__main__":
    main()
