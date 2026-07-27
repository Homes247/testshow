import os
import boto3
from botocore.exceptions import ClientError

R2_ENDPOINT_URL = os.getenv("R2_ENDPOINT_URL")
R2_ACCESS_KEY_ID = os.getenv("R2_ACCESS_KEY_ID")
R2_SECRET_ACCESS_KEY = os.getenv("R2_SECRET_ACCESS_KEY")
R2_BUCKET_NAME = os.getenv("R2_BUCKET_NAME")

_s3_client = boto3.client(
    "s3",
    endpoint_url=R2_ENDPOINT_URL,
    aws_access_key_id=R2_ACCESS_KEY_ID,
    aws_secret_access_key=R2_SECRET_ACCESS_KEY,
    region_name="auto",
)

# Maps the `doc_type` DB column to the existing bucket folder names
_DOC_TYPE_FOLDERS = {
    "doc": "Doc",
    "sheet": "Sheet",
    "slide": "Slide",
}


class DocumentStorage:
    @staticmethod
    def _key_for(owner_id: int, doc_id: str, doc_type: str = "doc") -> str:
        folder = _DOC_TYPE_FOLDERS.get(doc_type, "Doc")
        return f"Drive/{folder}/{owner_id}/{doc_id}.json"

    @staticmethod
    def relative_path(owner_id: int, doc_id: str, doc_type: str = "doc") -> str:
        """Returns the R2 object key. Kept as `relative_path` for naming
        continuity with the old local-disk implementation, but this is now
        the actual S3/R2 key stored in documents.file_path."""
        return DocumentStorage._key_for(owner_id, doc_id, doc_type)

    @staticmethod
    def save(owner_id: int, doc_id: str, content: str, doc_type: str = "doc") -> dict:
        import gzip
        key = DocumentStorage._key_for(owner_id, doc_id, doc_type)
        body_bytes = gzip.compress(content.encode("utf-8"), compresslevel=1)

        _s3_client.put_object(
            Bucket=R2_BUCKET_NAME,
            Key=key,
            Body=body_bytes,
            ContentType="application/json",
        )

        return {
            "relative_path": key,
            "size": len(body_bytes),
        }

    @staticmethod
    def load(owner_id: int, doc_id: str, doc_type: str = "doc", file_path: str = None) -> str:
        # Always prefer the DB's stored file_path as the source of truth —
        # only fall back to recomputing the key if file_path wasn't passed
        # (keeps backward compatibility with any call site not yet updated).
        key = file_path or DocumentStorage._key_for(owner_id, doc_id, doc_type)

        try:
            response = _s3_client.get_object(Bucket=R2_BUCKET_NAME, Key=key)
            body_bytes = response["Body"].read()
            import gzip
            try:
                return gzip.decompress(body_bytes).decode("utf-8")
            except Exception:
                return body_bytes.decode("utf-8")
        except ClientError as e:
            error_code = e.response.get("Error", {}).get("Code", "")
            if error_code in ("NoSuchKey", "404"):
                raise FileNotFoundError(f"Document not found in R2: {key}")
            raise

    @staticmethod
    def delete(owner_id: int, doc_id: str, doc_type: str = "doc", file_path: str = None) -> None:
        key = file_path or DocumentStorage._key_for(owner_id, doc_id, doc_type)
        try:
            _s3_client.delete_object(Bucket=R2_BUCKET_NAME, Key=key)
        except ClientError as e:
            # Deleting a non-existent key is not an error we need to surface —
            # log and continue so delete_document endpoint doesn't 500.
            print(f"R2 delete warning for key {key}: {e}")
