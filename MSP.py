from io import BytesIO

import openpyxl
import pandas as pd
import pymupdf
import streamlit as st


# -------------------------------------------------------
# OTOMATİK GİDER PUSULASI
# Devran Muhasebe
#
# Excel dosyasındaki kayıtları okur.
# Matbaadan basılmış yatay A4 koçan üzerine yazdırılmak
# üzere yalnızca bilgilerin bulunduğu PDF oluşturur.
# -------------------------------------------------------


# ===========================================================
# SAYFA ÖLÇÜLERİ
# A4 yatay PDF ölçüsü
# ===========================================================

SAYFA_GENISLIK = 842
SAYFA_YUKSEKLIK = 595

FONT_SIZE = 10


# ===========================================================
# SOL NÜSHA KOORDİNATLARI
# ===========================================================

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


# ===========================================================
# SAĞ NÜSHA KOORDİNATLARI
# ===========================================================

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
# YARDIMCI FONKSİYONLAR
# ===========================================================

def metin(deger):
    """
    Excel hücresindeki değeri temiz bir metne çevirir.
    Boş değerlerde 'nan' yazmasını engeller.
    """

    if deger is None or pd.isna(deger):
        return ""

    if isinstance(deger, float) and deger.is_integer():
        return str(int(deger))

    return str(deger).strip()


def tarih_formatla(deger):
    """
    Tarihi GG.AA.YYYY biçimine dönüştürür.
    """

    if deger is None or pd.isna(deger):
        return ""

    try:
        return pd.to_datetime(deger).strftime("%d.%m.%Y")

    except Exception:
        return metin(deger)


def para(deger):
    """
    Sayısal değeri Türkçe para biçimine dönüştürür.
    Örnek: 12500 -> 12.500,00
    """

    if deger is None or pd.isna(deger):
        return ""

    try:
        return (
            "{:,.2f}".format(float(deger))
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

    except (TypeError, ValueError):
        return ""


def yaz(sayfa, x, y, deger, boyut=FONT_SIZE):
    """
    Verilen metni PDF sayfasına yazar.
    """

    sayfa.insert_text(
        pymupdf.Point(x, y),
        metin(deger),
        fontsize=boyut,
        fontname="helv",
        color=(0, 0, 0),
    )


# ===========================================================
# EXCEL DOSYASINI OKU
# ===========================================================

def excel_verilerini_oku(excel_bytes):
    """
    Yüklenen Excel dosyasından kayıtları ve K1 tarihini okur.
    """

    excel_akisi = BytesIO(excel_bytes)

    df = pd.read_excel(
        excel_akisi,
        sheet_name="Sayfa1",
        engine="openpyxl",
    )

    excel_akisi.seek(0)

    kitap = openpyxl.load_workbook(
        excel_akisi,
        data_only=True,
        read_only=True,
    )

    try:
        sayfa = kitap["Sayfa1"]
        sabit_tarih = sayfa["K1"].value

    finally:
        kitap.close()

    gerekli_sutunlar = [
        "SATILAN CİNSİ",
        "İSİM",
        "TC",
        "ÖDEME ŞEKLİ",
        "TOPLAM TUTAR",
        "ALTIN GRAM",
        "BİRİM FİYAT",
    ]

    eksik_sutunlar = [
        sutun
        for sutun in gerekli_sutunlar
        if sutun not in df.columns
    ]

    if eksik_sutunlar:
        eksikler = ", ".join(eksik_sutunlar)

        raise ValueError(
            f"Excel dosyasında şu sütunlar bulunamadı: {eksikler}"
        )

    # İsmi boş veya 0 olan satırları çıkar
    isimler = df["İSİM"].fillna("").astype(str).str.strip()

    df = df[
        (isimler != "")
        & (isimler != "0")
        & (isimler.str.lower() != "nan")
    ].copy()

    if df.empty:
        raise ValueError(
            "Excel dosyasında PDF oluşturulacak geçerli kayıt bulunamadı."
        )

    return df, tarih_formatla(sabit_tarih)


# ===========================================================
# PDF OLUŞTUR
# ===========================================================

def pdf_olustur(excel_bytes):
    """
    Excel kayıtlarından PDF oluşturur ve PDF verisini byte olarak döndürür.
    """

    df, tarih = excel_verilerini_oku(excel_bytes)

    sonuc = pymupdf.open()

    try:
        for _, satir in df.iterrows():

            sayfa = sonuc.new_page(
                width=SAYFA_GENISLIK,
                height=SAYFA_YUKSEKLIK,
            )

            isim = metin(satir["İSİM"])
            tc = metin(satir["TC"])
            urun = metin(satir["SATILAN CİNSİ"])

            miktar = satir["ALTIN GRAM"]
            birim = satir["BİRİM FİYAT"]
            toplam = satir["GİDEN HAVALE TUTARI"]

            odeme_turu = metin(satir["ÖDEME ŞEKLİ"])

            # Aynı bilgiler hem sol hem sağ nüshaya yazılır
            for koordinatlar in (SOL, SAG):

                yaz(
                    sayfa,
                    *koordinatlar["tarih"],
                    tarih,
                )

                yaz(
                    sayfa,
                    *koordinatlar["isim"],
                    isim,
                )

                yaz(
                    sayfa,
                    *koordinatlar["tc"],
                    tc,
                )

                yaz(
                    sayfa,
                    *koordinatlar["urun"],
                    urun,
                )

                yaz(
                    sayfa,
                    *koordinatlar["miktar"],
                    para(miktar),
                )

                yaz(
                    sayfa,
                    *koordinatlar["birim"],
                    para(birim),
                )

                yaz(
                    sayfa,
                    *koordinatlar["toplam"],
                    para(toplam),
                )

                yaz(
                    sayfa,
                    *koordinatlar["odeme"],
                    odeme_turu,
                )

                yaz(
                    sayfa,
                    *koordinatlar["genel_toplam"],
                    para(toplam),
                )

        pdf_bytes = sonuc.tobytes(
            garbage=4,
            deflate=True,
        )

    finally:
        sonuc.close()

    return pdf_bytes, len(df)


# ===========================================================
# STREAMLIT SAYFA AYARLARI
# ===========================================================

st.set_page_config(
    page_title="Otomatik Gider Pusulası",
    page_icon="📄",
    layout="centered",
)


# ===========================================================
# SAYFA BAŞLIĞI
# ===========================================================

st.title("📄 Otomatik Gider Pusulası")

st.subheader("Devran Muhasebe")

st.write(
    "Excel dosyasını yükleyin. Program, matbaadan basılmış "
    "gider pusulası koçanına uygun PDF dosyasını oluşturacaktır."
)


# ===========================================================
# EXCEL YÜKLEME
# ===========================================================

excel_dosyasi = st.file_uploader(
    "Excel dosyasını seçin",
    type=["xlsx"],
    help="Excel dosyasının Sayfa1 sayfası kullanılacaktır.",
)


# ===========================================================
# PDF OLUŞTURMA
# ===========================================================

if excel_dosyasi is not None:

    st.success(
        f"Excel dosyası yüklendi: {excel_dosyasi.name}"
    )

    try:
        excel_bytes = excel_dosyasi.getvalue()

        pdf_bytes, kayit_sayisi = pdf_olustur(excel_bytes)

        st.info(
            f"Toplam {kayit_sayisi} kayıt bulundu. "
            f"PDF dosyası hazır."
        )

        st.download_button(
            label="⬇️ Gider Pusulalarını İndir",
            data=pdf_bytes,
            file_name="Gider_Pusulalari.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True,
        )

        st.warning(
            "Yazdırırken ölçek ayarını Gerçek Boyut veya %100 seçin. "
            "Sayfaya Sığdır seçeneğini kapatın."
        )

    except Exception as hata:

        st.error(
            f"PDF oluşturulamadı: {hata}"
        )

else:

    st.info(
        "PDF oluşturmak için yukarıdan bir Excel dosyası yükleyin."
    )