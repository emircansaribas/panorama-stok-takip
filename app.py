from flask import Flask, from zoneinfo import ZoneInfo, render_template, request, redirect, url_for, session

app = Flask(__name__)

import sqlite3


def kullanici_kontrol(kullanici, sifre):

    conn = sqlite3.connect("stok.db")

    cursor = conn.cursor()


    cursor.execute(
        """
        SELECT kullanici, yetki
        FROM kullanicilar
        WHERE kullanici=? AND sifre=?
        """,
        (
            kullanici,
            sifre
        )
    )


    sonuc = cursor.fetchone()


    conn.close()


    return sonuc

app.secret_key = "panorama_stok_2026"


KULLANICI = "admin"
SIFRE = "1234"



@app.route("/")
def home():
    return redirect(url_for("login"))



@app.route("/login", methods=["GET", "POST"])
def login():

    if request.method == "POST":

        kullanici = request.form["kullanici"]
        sifre = request.form["sifre"]


        sonuc = kullanici_kontrol(
            kullanici,
            sifre
        )


        if sonuc:

            session["kullanici"] = sonuc[0]

            session["yetki"] = sonuc[1]

            return redirect(url_for("dashboard"))



        return render_template(
            "login.html",
            hata="Kullanıcı adı veya şifre yanlış."
        )


    return render_template("login.html")



@app.route("/dashboard")
def dashboard():

    if "kullanici" not in session:
        return redirect(url_for("login"))

    import sqlite3

    conn = sqlite3.connect("stok.db")
    cursor = conn.cursor()

    # Toplam ürün
    cursor.execute("SELECT COUNT(*) FROM urunler")
    toplam_urun = cursor.fetchone()[0]

    # Toplam stok
    cursor.execute("SELECT SUM(stok) FROM urunler")
    toplam_stok = cursor.fetchone()[0] or 0

    # Eklenen stok
    cursor.execute("""
        SELECT SUM(miktar)
        FROM hareketler
        WHERE islem='Eklendi'
    """)
    eklenen = cursor.fetchone()[0] or 0

    # Harcanan stok
    cursor.execute("""
        SELECT SUM(miktar)
        FROM hareketler
        WHERE islem='Harcandı'
    """)
    harcanan = cursor.fetchone()[0] or 0

    # Kritik stok sayısı
    cursor.execute("""
        SELECT COUNT(*)
        FROM urunler
        WHERE stok <= kritik_seviye
    """)
    kritik = cursor.fetchone()[0]

    # Son 5 hareket
    cursor.execute("""
        SELECT
            urunler.urun_adi,
            hareketler.islem,
            hareketler.miktar,
            hareketler.tarih
        FROM hareketler
        JOIN urunler
            ON hareketler.urun_id = urunler.id
        ORDER BY hareketler.id DESC
        LIMIT 5
    """)
    son_hareketler = cursor.fetchall()

    # Kritik stoktaki ürünler
    cursor.execute("""
        SELECT
            urun_adi,
            stok,
            kritik_seviye
        FROM urunler
        WHERE stok <= kritik_seviye
    """)

    kritik_urunler = cursor.fetchall()

    print("KRİTİK ÜRÜNLER =", kritik_urunler)

    conn.close()

    return render_template(
        "dashboard.html",
        kullanici=session["kullanici"],
        toplam_urun=toplam_urun,
        toplam_stok=toplam_stok,
        eklenen=eklenen,
        harcanan=harcanan,
        kritik=kritik,
        son_hareketler=son_hareketler,
        kritik_urunler=kritik_urunler
    )


@app.route("/urun_ekle", methods=["GET", "POST"])
def urun_ekle():

    if "kullanici" not in session:
        return redirect(url_for("login"))

    if session.get("yetki") != "admin":
        return redirect(url_for("dashboard"))

    import sqlite3


    if request.method == "POST":

        urun_adi = request.form["urun_adi"]
        kategori = request.form["kategori"]
        stok = request.form["stok"]
        kritik = request.form["kritik"]


        conn = sqlite3.connect("stok.db")

        cursor = conn.cursor()


        cursor.execute("""
        INSERT INTO urunler
        (urun_adi, kategori, stok, kritik_seviye)

        VALUES (?, ?, ?, ?)
        """,
        (
            urun_adi,
            kategori,
            stok,
            kritik
        ))


        conn.commit()

        conn.close()


        return redirect(url_for("urun_ekle"))



    return render_template("urun_ekle.html")



@app.route("/stoklar")
def stoklar():

    if "kullanici" not in session:
        return redirect(url_for("login"))

    import sqlite3

    conn = sqlite3.connect("stok.db")
    cursor = conn.cursor()

    arama = request.args.get("arama", "")

    cursor.execute("""
        SELECT *
        FROM urunler
        WHERE urun_adi LIKE ?
        ORDER BY urun_adi
    """, (f"%{arama}%",))

    urunler = cursor.fetchall()

    conn.close()

    return render_template(
        "stoklar.html",
        urunler=urunler,
        arama=arama
    )

from flask import send_file
from openpyxl import Workbook
import io


@app.route("/excel_aktar")
def excel_aktar():

    if "kullanici" not in session:
        return redirect(url_for("login"))

    import sqlite3

    conn = sqlite3.connect("stok.db")
    cursor = conn.cursor()

    cursor.execute("""
        SELECT
            urun_adi,
            kategori,
            stok,
            kritik_seviye
        FROM urunler
        ORDER BY urun_adi
    """)

    urunler = cursor.fetchall()

    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Toplam Stok"

    ws.append([
        "Ürün",
        "Kategori",
        "Stok",
        "Kritik Seviye"
    ])

    for urun in urunler:
        ws.append(urun)

    dosya = io.BytesIO()
    wb.save(dosya)
    dosya.seek(0)

    return send_file(
        dosya,
        as_attachment=True,
        download_name="Toplam_Stok.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
@app.route("/pdf_rapor")
def pdf_rapor():

    if "kullanici" not in session:
        return redirect(url_for("login"))

    import sqlite3
    import io

    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.pagesizes import A4


    conn = sqlite3.connect("stok.db")
    cursor = conn.cursor()


    cursor.execute("""
        SELECT
            urun_adi,
            kategori,
            stok,
            kritik_seviye
        FROM urunler
        ORDER BY urun_adi
    """)


    urunler = cursor.fetchall()

    conn.close()


    dosya = io.BytesIO()


    pdf = SimpleDocTemplate(
        dosya,
        pagesize=A4
    )


    elemanlar = []


    stiller = getSampleStyleSheet()


    baslik = Paragraph(
        "PANORAMA STOK TAKIP RAPORU",
        stiller["Title"]
    )


    elemanlar.append(baslik)


    tablo = [
        [
            "Ürün",
            "Kategori",
            "Stok",
            "Kritik"
        ]
    ]


    for urun in urunler:
        tablo.append(urun)


    tablo_obj = Table(tablo)


    tablo_obj.setStyle(
        TableStyle([
            ("GRID", (0,0), (-1,-1), 1, None)
        ])
    )


    elemanlar.append(tablo_obj)


    pdf.build(elemanlar)


    dosya.seek(0)


    return send_file(
        dosya,
        as_attachment=True,
        download_name="Stok_Raporu.pdf",
        mimetype="application/pdf"
    )



@app.route("/stok_ekle", methods=["GET", "POST"])
def stok_ekle():

    if "kullanici" not in session:
        return redirect(url_for("login"))

    if session.get("yetki") != "admin":
        return redirect(url_for("dashboard"))

    import sqlite3
    from datetime import datetime


    conn = sqlite3.connect("stok.db")

    cursor = conn.cursor()


    if request.method == "POST":

        urun_id = request.form["urun_id"]

        miktar = int(request.form["miktar"])


        cursor.execute(
            "UPDATE urunler SET stok = stok + ? WHERE id = ?",
            (miktar, urun_id)
        )


        cursor.execute("""
        INSERT INTO hareketler
        (urun_id, islem, miktar, tarih, kullanici)

        VALUES (?, ?, ?, ?, ?)
        """,
        (
            urun_id,
            "Eklendi",
            miktar,
            datetime.now(ZoneInfo("Europe/Istanbul")).strftime("%d.%m.%Y %H:%M")
            session["kullanici"]
        ))


        conn.commit()

        conn.close()


        return redirect(url_for("stok_ekle"))



    cursor.execute("SELECT * FROM urunler")

    urunler = cursor.fetchall()


    conn.close()


    return render_template(
        "stok_ekle.html",
        urunler=urunler
    )


@app.route("/stok_harcama", methods=["GET", "POST"])
def stok_harcama():

    if "kullanici" not in session:
        return redirect(url_for("login"))

    if session.get("yetki") != "admin":
        return redirect(url_for("dashboard"))

    import sqlite3
    from datetime import datetime


    conn = sqlite3.connect("stok.db")

    cursor = conn.cursor()



    if request.method == "POST":

        urun_id = request.form["urun_id"]

        miktar = int(request.form["miktar"])


        cursor.execute(
            "SELECT stok FROM urunler WHERE id=?",
            (urun_id,)
        )

        mevcut = cursor.fetchone()[0]


        if miktar <= mevcut:


            cursor.execute(
                """
                UPDATE urunler

                SET stok = stok - ?

                WHERE id = ?
                """,
                (
                    miktar,
                    urun_id
                )
            )


            cursor.execute("""
            INSERT INTO hareketler
            (urun_id, islem, miktar, tarih, kullanici)

            VALUES (?, ?, ?, ?, ?)
            """,
            (
                urun_id,
                "Harcandı",
                miktar,
                datetime.now().strftime("%d.%m.%Y %H:%M"),
                session["kullanici"]
            ))


            conn.commit()


        conn.close()


        return redirect(url_for("stok_harcama"))



    cursor.execute("SELECT * FROM urunler")


    urunler = cursor.fetchall()


    conn.close()


    return render_template(
        "stok_harcama.html",
        urunler=urunler
    )



@app.route("/hareketler")
def hareketler():

    if "kullanici" not in session:
        return redirect(url_for("login"))


    import sqlite3


    conn = sqlite3.connect("stok.db")

    cursor = conn.cursor()


    cursor.execute("""
    SELECT

    hareketler.id,
    hareketler.urun_id,
    hareketler.islem,
    hareketler.miktar,
    hareketler.tarih,
    hareketler.kullanici,
    urunler.urun_adi


    FROM hareketler


    JOIN urunler

    ON hareketler.urun_id = urunler.id


    ORDER BY hareketler.id DESC

    """)


    hareketler = cursor.fetchall()


    conn.close()


    return render_template(
        "hareketler.html",
        hareketler=hareketler
    )



@app.route("/eklenenler")
def eklenenler():

    if "kullanici" not in session:
        return redirect(url_for("login"))


    import sqlite3


    conn = sqlite3.connect("stok.db")

    cursor = conn.cursor()


    cursor.execute("""
    SELECT

    hareketler.*,
    urunler.urun_adi


    FROM hareketler


    JOIN urunler

    ON hareketler.urun_id = urunler.id


    WHERE hareketler.islem = 'Eklendi'


    ORDER BY hareketler.id DESC

    """)


    hareketler = cursor.fetchall()


    conn.close()


    return render_template(
        "eklenenler.html",
        hareketler=hareketler
    )



@app.route("/harcananlar")
def harcananlar():

    if "kullanici" not in session:
        return redirect(url_for("login"))

    import sqlite3

    conn = sqlite3.connect("stok.db")
    cursor = conn.cursor()

    cursor.execute("""
        SELECT
            hareketler.id,
            urunler.urun_adi,
            hareketler.miktar,
            hareketler.tarih,
            hareketler.kullanici

        FROM hareketler

        JOIN urunler
        ON hareketler.urun_id = urunler.id

        WHERE hareketler.islem='Harcandı'

        ORDER BY hareketler.id DESC
    """)

    hareketler = cursor.fetchall()

    conn.close()

    return render_template(
        "harcananlar.html",
        hareketler=hareketler
    )
    
   
@app.route("/urun_duzenle/<int:id>", methods=["GET", "POST"])
def urun_duzenle(id):

    if "kullanici" not in session:
        return redirect(url_for("login"))

    if session.get("yetki") != "admin":
        return redirect(url_for("dashboard"))

    import sqlite3

    conn = sqlite3.connect("stok.db")
    cursor = conn.cursor()

    if request.method == "POST":

        urun_adi = request.form["urun_adi"]
        kategori = request.form["kategori"]
        stok = request.form["stok"]
        kritik = request.form["kritik"]

        cursor.execute("""
            UPDATE urunler
            SET urun_adi=?,
                kategori=?,
                stok=?,
                kritik_seviye=?
            WHERE id=?
        """,
        (
            urun_adi,
            kategori,
            stok,
            kritik,
            id
        ))

        conn.commit()
        conn.close()

        return redirect(url_for("stoklar"))

    cursor.execute(
        "SELECT * FROM urunler WHERE id=?",
        (id,)
    )

    urun = cursor.fetchone()

    conn.close()

    return render_template(
        "urun_duzenle.html",
        urun=urun
    )


@app.route("/urun_sil/<int:id>")
def urun_sil(id):

    if "kullanici" not in session:
        return redirect(url_for("login"))

    if session.get("yetki") != "admin":
        return redirect(url_for("dashboard"))

    import sqlite3

    conn = sqlite3.connect("stok.db")
    cursor = conn.cursor()

    cursor.execute(
        "DELETE FROM urunler WHERE id=?",
        (id,)
    )

    conn.commit()
    conn.close()

    return redirect(url_for("stoklar"))


@app.route("/logout")
def logout():

    session.clear()

    return redirect(url_for("login"))


if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )