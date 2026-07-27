import os
import asyncio
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from app.database import SessionLocal
from app.models.document import Document
from app.lib.document_storage import _s3_client, R2_BUCKET_NAME, _DOC_TYPE_FOLDERS
from sqlalchemy import select

LOCAL_STORAGE_BASE = os.path.join(os.getcwd(), "storage", "documents")


async def migrate():
    migrated = 0
    skipped = 0
    failed = 0

    async with SessionLocal() as db:
        result = await db.execute(
            select(Document).where(Document.file_path.isnot(None))
        )
        docs = result.scalars().all()

        for doc in docs:
            # Check if it looks like it's already an R2 path (starts with Drive/)
            if doc.file_path.startswith("Drive/"):
                print(f"[SKIP] Already an R2 path for {doc.id}: {doc.file_path}")
                skipped += 1
                continue
                
            local_path = os.path.join(LOCAL_STORAGE_BASE, doc.file_path)

            if not os.path.exists(local_path):
                print(f"[SKIP] No local file for {doc.id}: {local_path}")
                skipped += 1
                continue

            try:
                with open(local_path, "r", encoding="utf-8") as f:
                    content = f.read()

                folder = _DOC_TYPE_FOLDERS.get(doc.doc_type, "Doc")
                new_key = f"Drive/{folder}/{doc.owner_id}/{doc.id}.json"

                _s3_client.put_object(
                    Bucket=R2_BUCKET_NAME,
                    Key=new_key,
                    Body=content.encode("utf-8"),
                    ContentType="application/json",
                )

                doc.file_path = new_key
                doc.file_size = len(content.encode("utf-8"))
                migrated += 1
                print(f"[OK] {doc.id} -> {new_key}")

            except Exception as e:
                failed += 1
                print(f"[FAIL] {doc.id}: {e}")

        await db.commit()

    print(f"\nDone. Migrated: {migrated}, Skipped: {skipped}, Failed: {failed}")


if __name__ == "__main__":
    asyncio.run(migrate())
