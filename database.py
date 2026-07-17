import sqlite3


def veritabani_olustur():

    conn = sqlite3.connect("stok.db")

    cursor = conn.cursor()


    cursor.execute("""
    CREATE TABLE IF NOT EXISTS urunler (

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        urun_adi TEXT NOT NULL,

        kategori TEXT NOT NULL,

        stok INTEGER DEFAULT 0,

        kritik_seviye INTEGER DEFAULT 5

    )
    """)


    cursor.execute("""
    CREATE TABLE IF NOT EXISTS hareketler (

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        urun_id INTEGER,

        islem TEXT,

        miktar INTEGER,

        tarih TEXT,

        kullanici TEXT

    )
    """)


    conn.commit()

    conn.close()



veritabani_olustur()