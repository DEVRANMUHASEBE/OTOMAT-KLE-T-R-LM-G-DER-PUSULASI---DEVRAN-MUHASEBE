import os
import fitz
import pandas as pd
import openpyxl

from tkinter import *
from tkinter import filedialog, messagebox


# -------------------------------------------------------
# GİDER PUSULASI DOLDURMA PROGRAMI
# Yasin Ceylan Kuyumculuk
#
# PDF şablonu kullanılmaz.
# Yalnızca yazılar boş PDF üzerine eklenir.
# Oluşturulan PDF, matbaa koçanına yazdırılır.
# -------------------------------------------------------

FONT_SIZE = 10

# A4 yatay ölçüsü — PDF point değeri
SAYFA_GENISLIK = 842
SAYFA_YUKSEKLIK = 595


# ===========================
# SOL NÜSHA KOORDİNATLARI
# ===========================

SOL = {
    "tarih": (300, 79),

    "isim": (150, 152),
    "tc": (150, 165),

    "urun": (55, 255),
    "miktar": (225, 255),
    "birim": (300, 255),
    "toplam": (350, 255),

    "odeme": (105, 462),
    "genel_toplam": (105, 474),
}


# ===========================
# SAĞ NÜSHA KOORDİNATLARI
# ===========================

SAG = {
    "tarih": (701, 79),

    "isim": (541, 152),
    "tc": (541, 165),

    "urun": (456, 255),
    "miktar": (626, 255),
    "birim": (706, 255),
    "toplam": (766, 255),

    "odeme": (506, 462),
    "genel_toplam": (506, 474),
}


# ===========================================================
# PDF ÜZERİNE YAZI YAZ
# ===========================================================

def yaz(page, x, y, text, boyut=FONT_SIZE):

    if text is None:
        text = ""

    page.insert_text(
        fitz.Point(x, y),
        str(text),
        fontsize=boyut,
        fontname="helv",
        color=(0, 0, 0),
    )


# ===========================================================
# BOŞ DEĞERLERİ TEMİZLE
# ===========================================================

def metin(deger):

    if deger is None or pd.isna(deger):
        return ""

    # Excel'den 12345678901.0 şeklinde gelen değerleri düzeltir
    if isinstance(deger, float) and deger.is_integer():
        return str(int(deger))

    return str(deger).strip()


# ===========================================================
# TARİH FORMATI
# ===========================================================

def tarih_formatla(deger):

    try:
        if deger is None or pd.isna(deger):
            return ""

        return pd.to_datetime(deger).strftime("%d.%m.%Y")

    except Exception:
        return metin(deger)


# ===========================================================
# PARA FORMATI
# ===========================================================

def para(deger):

    try:
        if deger is None or pd.isna(deger):
            return ""

        return (
            "{:,.2f}".format(float(deger))
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

    except Exception:
        return ""


# ===========================================================
# PDF OLUŞTUR
# ===========================================================

def pdf_olustur(excel_dosyasi, cikti):

    # Excel tablosunu oku
    df = pd.read_excel(
        excel_dosyasi,
        sheet_name="Sayfa1",
        engine="openpyxl"
    )

    # Sabit tarihi K1 hücresinden al
    excel_kitap = openpyxl.load_workbook(
        excel_dosyasi,
        data_only=True
    )

    excel_sayfa = excel_kitap["Sayfa1"]
    sabit_tarih = excel_sayfa["K1"].value
    excel_kitap.close()

    tarih = tarih_formatla(sabit_tarih)

    # İsmi boş veya 0 olan satırları alma
    df = df[
        df["İSİM"].notna()
        & (df["İSİM"].astype(str).str.strip() != "")
        & (df["İSİM"].astype(str).str.strip() != "0")
    ].copy()

    if df.empty:
        raise ValueError("Excel dosyasında yazdırılacak kayıt bulunamadı.")

    sonuc = fitz.open()

    for _, row in df.iterrows():

        # Her Excel kaydı için boş yatay A4 sayfa oluştur
        sayfa = sonuc.new_page(
            width=SAYFA_GENISLIK,
            height=SAYFA_YUKSEKLIK
        )

        isim = metin(row["İSİM"])
        tc = metin(row["TC"])
        urun = metin(row["SATILAN CİNSİ"])

        miktar = row["ALTIN GRAM"]
        birim = row["BİRİM FİYAT"]
        toplam = row["GİDEN HAVALE TUTARI"]

        odeme_turu = metin(row["ÖDEME ŞEKLİ"])

        # Aynı bilgileri sol ve sağ nüshaya yaz
        for K in (SOL, SAG):

            yaz(sayfa, *K["tarih"], tarih)

            yaz(sayfa, *K["isim"], isim)
            yaz(sayfa, *K["tc"], tc)

            yaz(sayfa, *K["urun"], urun)
            yaz(sayfa, *K["miktar"], para(miktar))
            yaz(sayfa, *K["birim"], para(birim))
            yaz(sayfa, *K["toplam"], para(toplam))

            yaz(sayfa, *K["odeme"], odeme_turu)
            yaz(sayfa, *K["genel_toplam"], para(toplam))

    sonuc.save(
        cikti,
        garbage=4,
        deflate=True
    )

    sonuc.close()


# ===========================================================
# EXCEL DOSYASI SEÇ
# ===========================================================

def excel_sec():

    dosya = filedialog.askopenfilename(
        title="Excel Dosyasını Seç",
        filetypes=[
            ("Excel Dosyası", "*.xlsx *.xls")
        ]
    )

    if dosya:
        excel_var.set(dosya)


# ===========================================================
# PDF OLUŞTUR VE KAYDET
# ===========================================================

def pdf_uret():

    if excel_var.get().strip() == "":
        messagebox.showerror(
            "Hata",
            "Excel dosyasını seçiniz."
        )
        return

    cikti = filedialog.asksaveasfilename(
        title="Gider Pusulalarını Kaydet",
        defaultextension=".pdf",
        filetypes=[
            ("PDF Dosyası", "*.pdf")
        ],
        initialfile="Gider_Pusulalari.pdf"
    )

    if cikti == "":
        return

    try:

        pdf_olustur(
            excel_var.get(),
            cikti
        )

        messagebox.showinfo(
            "Başarılı",
            "PDF başarıyla oluşturuldu ve kaydedildi."
        )

    except Exception as hata:

        messagebox.showerror(
            "Hata",
            str(hata)
        )


# ===========================================================
# KAYITLI PDF AÇ
# ===========================================================

def pdf_ac():

    dosya = filedialog.askopenfilename(
        title="Oluşturulan PDF'yi Seç",
        filetypes=[
            ("PDF Dosyası", "*.pdf")
        ]
    )

    if dosya:

        try:
            os.startfile(dosya)

        except Exception:
            messagebox.showerror(
                "Hata",
                "PDF açılamadı."
            )


# ===========================================================
# PROGRAMI KAPAT
# ===========================================================

def cikis():
    pencere.destroy()


# ===========================================================
# TKINTER ARAYÜZÜ
# ===========================================================

pencere = Tk()

pencere.title("Yasin Ceylan Kuyumculuk")
pencere.geometry("620x250")
pencere.resizable(False, False)

excel_var = StringVar()


Label(
    pencere,
    text="Excel Dosyası",
    font=("Arial", 10, "bold")
).place(x=20, y=35)


Entry(
    pencere,
    textvariable=excel_var,
    width=60
).place(x=130, y=38)


Button(
    pencere,
    text="Seç",
    width=10,
    command=excel_sec
).place(x=510, y=34)


Button(
    pencere,
    text="PDF OLUŞTUR",
    bg="green",
    fg="white",
    font=("Arial", 12, "bold"),
    width=20,
    height=2,
    command=pdf_uret
).place(x=195, y=100)


Button(
    pencere,
    text="PDF AÇ",
    width=15,
    command=pdf_ac
).place(x=110, y=195)


Button(
    pencere,
    text="ÇIKIŞ",
    width=15,
    bg="red",
    fg="white",
    command=cikis
).place(x=350, y=195)


# ===========================================================
# PROGRAMI BAŞLAT
# ===========================================================

pencere.mainloop()