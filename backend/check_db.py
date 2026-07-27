import pymysql
conn = pymysql.connect(host='74.225.249.46', port=3307, user='chat_bot', password='password', database='vmail')
cursor = conn.cursor(pymysql.cursors.DictCursor)
cursor.execute('SELECT id, title, doc_type, file_path FROM documents WHERE doc_type="sheet" ORDER BY updated_at DESC LIMIT 2')
rows = cursor.fetchall()
for row in rows:
    print(f"ID: {row['id']}, Path: {row['file_path']}, Content len: {len(row.get('content', '') or '')}")
conn.close()
