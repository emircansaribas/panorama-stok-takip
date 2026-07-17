import sqlite3


conn = sqlite3.connect("stok.db")

cursor = conn.cursor()


cursor.execute("""
CREATE TABLE IF NOT EXISTS kullanicilar (

    id INTEGER PRIMARY KEY AUTOINCREMENT,

    kullanici TEXT UNIQUE,

    sifre TEXT,

    yetki TEXT

)
""")


cursor.execute("""
INSERT OR IGNORE INTO kullanicilar
(kullanici, sifre, yetki)

VALUES (?, ?, ?)
""",
(
    "admin",
    "1234",
    "admin"
))


cursor.execute("""
INSERT OR IGNORE INTO kullanicilar
(kullanici, sifre, yetki)

VALUES (?, ?, ?)
""",
(
    "personel",
    "1234",
    "personel"
))


conn.commit()

conn.close()


print("Kullanıcı sistemi hazır")