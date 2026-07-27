import re

path = "backend/app/api/documents.py"
with open(path, "r", encoding="utf-8") as f:
    content = f.read()

# Fix create_document
content = re.sub(
    r"    doc = Document\(\n        title=body\.title,\n        doc_type=body\.doc_type,\n        content=_default_content\(body\.doc_type, body\.title\),\n        owner_id=current_user\.id,\n    \)",
    "    content_val = _default_content(body.doc_type, body.title)\n    doc = Document(\n        title=body.title,\n        doc_type=body.doc_type,\n        owner_id=current_user.id,\n    )",
    content,
    flags=re.MULTILINE
)

content = re.sub(
    r"    try:\n        storage_res = DocumentStorage\.save\(doc\.owner_id, doc\.id, doc\.content\)",
    "    try:\n        storage_res = DocumentStorage.save(doc.owner_id, doc.id, content_val)",
    content,
    flags=re.MULTILINE
)

# Fix import_document
content = re.sub(
    r"    doc = Document\(\n        title=title,\n        doc_type=doc_type,\n        content=content_json,\n        owner_id=current_user\.id,\n    \)",
    "    doc = Document(\n        title=title,\n        doc_type=doc_type,\n        owner_id=current_user.id,\n    )",
    content,
    flags=re.MULTILINE
)

content = re.sub(
    r"    try:\n        storage_res = DocumentStorage\.save\(doc\.owner_id, doc\.id, doc\.content\)",
    "    try:\n        storage_res = DocumentStorage.save(doc.owner_id, doc.id, content_json)",
    content,
    flags=re.MULTILINE
)

# Fix get_document
content = re.sub(
    r"    content = doc\.content",
    "    content = \"{}\"",
    content,
    flags=re.MULTILINE
)

# Fix update_document
content = re.sub(
    r"    if body\.content is not None:\n        doc\.content = body\.content\n        try:",
    "    if body.content is not None:\n        try:",
    content,
    flags=re.MULTILINE
)

with open(path, "w", encoding="utf-8") as f:
    f.write(content)
