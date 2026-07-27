import sqlite3
conn = sqlite3.connect('office_suite.db')
print(conn.execute('SELECT email, phone FROM user').fetchall())
