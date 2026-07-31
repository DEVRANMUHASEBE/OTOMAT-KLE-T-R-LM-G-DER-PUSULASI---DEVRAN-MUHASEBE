from io import BytesIO
from pathlib import Path
from datetime import date
import unicodedata

import openpyxl
import pandas as pd
import pymupdf
import streamlit as st

# ==================================================
# 🔐 KULLANICI GİRİŞ SİSTEMİ
# ==================================================

if "giris_yapildi" not in st.session_state:
    st.session_state.giris_yapildi = False

if "aktif_kullanici" not in st.session_state:
    st.session_state.aktif_kullanici = ""

if "giris_hatasi" not in st.session_state:
    st.session_state.giris_hatasi = False


def giris_yap():
    kullanici_adi = st.session_state.get("login_kullanici", "").strip()
    sifre = st.session_state.get("login_sifre", "")

    try:
        kullanicilar = st.secrets["kullanicilar"]

        if (
            kullanici_adi in kullanicilar
            and kullanicilar[kullanici_adi] == sifre
        ):
            st.session_state.giris_yapildi = True
            st.session_state.aktif_kullanici = kullanici_adi
            st.session_state.giris_hatasi = False
            st.rerun()

        else:
            st.session_state.giris_hatasi = True

    except Exception:
        st.error("⚠️ Kullanıcı ayarları okunamadı.")
        st.stop()


# ==================================================
# 🔐 GİRİŞ EKRANI
# ==================================================

if not st.session_state.giris_yapildi:

    st.markdown(
    """
    <div style="text-align:center; margin-top:80px; margin-bottom:30px;">
        <h1>🔐 Devran Muhasebe</h1>
        <p style="font-size:18px; color:#888;">Gider Pusulası Sistemine Giriş</p>
    </div>
    """,
    unsafe_allow_html=True
)

    col1, col2, col3 = st.columns([1, 2, 1])

    with col2:

        st.text_input(
            "👤 Kullanıcı Adı",
            key="login_kullanici",
            placeholder="Kullanıcı adınızı giriniz"
        )

        st.text_input(
            "🔑 Şifre",
            type="password",
            key="login_sifre",
            placeholder="Şifrenizi giriniz"
        )

        if st.button(
            "🔓 Giriş Yap",
            use_container_width=True,
            type="primary"
        ):
            giris_yap()

        if st.session_state.giris_hatasi:
            st.error("❌ Kullanıcı adı veya şifre hatalı.")

    # KRİTİK:
    # Giriş yapılmadan aşağıdaki MSP.py kodları çalışmaz.
    st.stop()
    
# -------------------------------------------------------
# OTOMATİK VE MANUEL GİDER PUSULASI
# Devran Muhasebe
# -------------------------------------------------------


# ===========================================================
# FONT AYARLARI
# ===========================================================

PROJE_KLASORU = Path(__file__).resolve().parent
FONT_DOSYASI = PROJE_KLASORU / "fonts" / "DejaVuSans.ttf"
FONT_ADI = "DevranFont"


if not FONT_DOSYASI.is_file():
    st.error(
        "DejaVuSans.ttf bulunamadı. "
        "Dosyanın fonts/DejaVuSans.ttf yolunda olduğundan emin olun."
    )
    st.stop()


# ===========================================================
# SAYFA ÖLÇÜLERİ
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

def unicode_duzelt(deger):
    """
    Türkçe karakterleri Unicode NFC biçimine dönüştürür.
    """

    return unicodedata.normalize("NFC", str(deger))


def metin(deger):
    """
    Değeri temiz metne dönüştürür.
    """

    if deger is None or pd.isna(deger):
        return ""

    if isinstance(deger, float) and deger.is_integer():
        return str(int(deger))

    return unicode_duzelt(deger).strip()


def tarih_formatla(deger):
    """
    Tarihi GG.AA.YYYY biçiminde döndürür.
    """

    if deger is None or pd.isna(deger):
        return ""

    try:
        return pd.to_datetime(deger).strftime("%d.%m.%Y")

    except Exception:
        return metin(deger)


def para(deger):
    """
    Sayıyı Türkçe para biçimine dönüştürür.

    12500 -> 12.500,00
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
    PDF üzerine Türkçe karakter destekli metin yazar.
    """

    yazilacak_metin = metin(deger)

    if not yazilacak_metin:
        return

    sayfa.insert_text(
        pymupdf.Point(x, y),
        yazilacak_metin,
        fontsize=boyut,
        fontname=FONT_ADI,
        fontfile=str(FONT_DOSYASI),
        color=(0, 0, 0),
        overlay=True,
    )


# ===========================================================
# EXCEL DOSYASINI OKU
# ===========================================================

def excel_verilerini_oku(excel_bytes):
    """
    Excel dosyasından kayıtları ve L1 hücresindeki tarihi okur.
    """

    excel_akisi = BytesIO(excel_bytes)

    df = pd.read_excel(
        excel_akisi,
        sheet_name="Sayfa1",
        engine="openpyxl",
    )

    df.columns = [
        unicode_duzelt(sutun).strip()
        for sutun in df.columns
    ]

    excel_akisi.seek(0)

    kitap = openpyxl.load_workbook(
        excel_akisi,
        data_only=True,
        read_only=True,
    )

    try:
        sayfa = kitap["Sayfa1"]
        sabit_tarih = sayfa["L1"].value

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
        raise ValueError(
            "Excel dosyasında şu sütunlar bulunamadı: "
            + ", ".join(eksik_sutunlar)
        )

    isimler = (
        df["İSİM"]
        .fillna("")
        .astype(str)
        .map(unicode_duzelt)
        .str.strip()
    )

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
# PDF OLUŞTURMA
# ===========================================================

def kayitlari_pdf_yap(df, tarih):
    """
    DataFrame içindeki kayıtları PDF dosyasına dönüştürür.
    Excel ve manuel giriş aynı fonksiyonu kullanır.
    """

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
            toplam = satir["TOPLAM TUTAR"]

            odeme_turu = metin(satir["ÖDEME ŞEKLİ"])

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

    return pdf_bytes


def excel_pdf_olustur(excel_bytes):
    """
    Excel dosyasındaki kayıtları PDF'ye dönüştürür.
    """

    df, tarih = excel_verilerini_oku(excel_bytes)

    pdf_bytes = kayitlari_pdf_yap(
        df=df,
        tarih=tarih,
    )

    return pdf_bytes, len(df)


def manuel_pdf_olustur(
    tarih,
    isim,
    tc,
    urun,
    miktar,
    birim_fiyat,
    toplam_tutar,
    odeme_sekli,
):
    """
    Manuel girilen tek kaydı PDF'ye dönüştürür.
    """

    manuel_df = pd.DataFrame(
        [
            {
                "SATILAN CİNSİ": urun,
                "İSİM": isim,
                "TC": tc,
                "ÖDEME ŞEKLİ": odeme_sekli,
                "TOPLAM TUTAR": toplam_tutar,
                "ALTIN GRAM": miktar,
                "BİRİM FİYAT": birim_fiyat,
            }
        ]
    )

    return kayitlari_pdf_yap(
        df=manuel_df,
        tarih=tarih_formatla(tarih),
    )

# ===========================================================
# MANUEL HESAPLAMA
# ===========================================================

def toplam_tutari_hesapla():
    miktar = st.session_state.get("manuel_miktar", 0.0)
    birim = st.session_state.get("manuel_birim_fiyat", 0.0)

    if miktar > 0:
        st.session_state["manuel_toplam"] = round(
            miktar * birim,
            2,
        )


def birim_fiyati_hesapla():
    miktar = st.session_state.get("manuel_miktar", 0.0)
    toplam = st.session_state.get("manuel_toplam", 0.0)

    if miktar > 0:
        st.session_state["manuel_birim_fiyat"] = round(
            toplam / miktar,
            2,
        )


def miktar_degisince_hesapla():

    miktar = st.session_state.get("manuel_miktar", 0.0)
    birim = st.session_state.get("manuel_birim_fiyat", 0.0)
    toplam = st.session_state.get("manuel_toplam", 0.0)

    if miktar <= 0:
        return

    if birim > 0:
        st.session_state["manuel_toplam"] = round(
            miktar * birim,
            2,
        )

    elif toplam > 0:
        st.session_state["manuel_birim_fiyat"] = round(
            toplam / miktar,
            2,
        )
# ===========================================================
# STREAMLIT AYARLARI
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
st.subheader("Devran Mâli Müşavirlik")

st.write(
    "Toplu Excel dosyası yükleyebilir veya gider pusulasını "
    "manuel olarak hazırlayabilirsiniz."
)


# ===========================================================
# SEKMELER
# ===========================================================

excel_sekmesi, manuel_sekmesi = st.tabs(
    [
        "📊 Toplu Excel",
        "✍️ Manuel Giriş",
    ]
)


# ===========================================================
# TOPLU EXCEL SEKMESİ
# ===========================================================

with excel_sekmesi:

    st.markdown("### Toplu Excel işlemi")

    st.write(
        "Excel dosyasını yükleyin. Program tüm geçerli kayıtları "
        "tek bir PDF dosyasında oluşturacaktır."
    )

    excel_dosyasi = st.file_uploader(
        "Excel dosyasını seçin",
        type=["xlsx"],
        help="Excel dosyasının Sayfa1 sayfası kullanılacaktır.",
        key="excel_yukleme",
    )

    if excel_dosyasi is not None:

        st.success(
            f"Excel dosyası yüklendi: {excel_dosyasi.name}"
        )

        try:
            excel_bytes = excel_dosyasi.getvalue()

            pdf_bytes, kayit_sayisi = excel_pdf_olustur(
                excel_bytes
            )

            st.info(
                f"Toplam {kayit_sayisi} kayıt bulundu. "
                "PDF dosyası hazır."
            )

            st.download_button(
                label="⬇️ Toplu Gider Pusulalarını İndir",
                data=pdf_bytes,
                file_name="Toplu_Gider_Pusulalari.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True,
                key="toplu_pdf_indir",
            )

        except Exception as hata:
            st.error(
                f"PDF oluşturulamadı: {hata}"
            )

    else:
        st.info(
            "PDF oluşturmak için yukarıdan bir Excel dosyası yükleyin."
        )


# ===========================================================
# MANUEL GİRİŞ SEKMESİ
# ===========================================================

with manuel_sekmesi:

    st.markdown("### ✍️ Manuel gider pusulası")

    st.write(
        "Bilgileri doldurun ve tek kişilik gider pusulasını oluşturun."
    )

    st.info(
        "Miktar ve birim fiyat girildiğinde toplam tutar; "
        "miktar ve toplam tutar girildiğinde birim fiyat "
        "otomatik hesaplanır."
    )

    # -------------------------------------------------------
    # TARİH
    # -------------------------------------------------------

    manuel_tarih = st.date_input(
        "Tarih",
        value=date.today(),
        format="DD.MM.YYYY",
        key="manuel_tarih",
    )

    # -------------------------------------------------------
    # SATILAN CİNSİ
    # -------------------------------------------------------

    manuel_urun = st.text_input(
        "Satılan cinsi",
        placeholder="Örneğin: Çeyrek Altın",
        key="manuel_urun",
    )

    # -------------------------------------------------------
    # TC VE İSİM
    # -------------------------------------------------------

    kimlik_sutunu, isim_sutunu = st.columns(2)

    with kimlik_sutunu:

        manuel_tc = st.text_input(
            "T.C. kimlik numarası",
            max_chars=11,
            placeholder="11 haneli T.C. numarası",
            key="manuel_tc",
        )

    with isim_sutunu:

        manuel_isim = st.text_input(
            "İsim soyisim",
            placeholder="Ad Soyad",
            key="manuel_isim",
        )

    # -------------------------------------------------------
    # TUTAR BİLGİLERİ
    # -------------------------------------------------------

    st.markdown("#### Tutar bilgileri")

    miktar_sutunu, birim_sutunu, toplam_sutunu = st.columns(3)

    with miktar_sutunu:

        manuel_miktar = st.number_input(
            "Miktar",
            min_value=0.0,
            value=0.0,
            step=0.01,
            format="%.2f",
            key="manuel_miktar",
            on_change=miktar_degisince_hesapla,
            help=(
                "Miktar değiştirildiğinde, dolu olan tutar "
                "alanına göre yeniden hesaplama yapılır."
            ),
        )

    with birim_sutunu:

        manuel_birim_fiyat = st.number_input(
            "Birim fiyat",
            min_value=0.0,
            value=0.0,
            step=0.01,
            format="%.2f",
            key="manuel_birim_fiyat",
            on_change=toplam_tutari_hesapla,
            help=(
                "Miktar ile birlikte girildiğinde toplam tutar "
                "otomatik hesaplanır."
            ),
        )

    with toplam_sutunu:

        manuel_toplam = st.number_input(
            "Toplam tutar",
            min_value=0.0,
            value=0.0,
            step=0.01,
            format="%.2f",
            key="manuel_toplam",
            on_change=birim_fiyati_hesapla,
            help=(
                "Miktar ile birlikte girildiğinde birim fiyat "
                "otomatik hesaplanır."
            ),
        )

    # -------------------------------------------------------
    # ÖDEME ŞEKLİ
    # -------------------------------------------------------

    manuel_odeme = st.text_input(
        "Ödeme şekli",
        placeholder="Örneğin: Nakit, banka, havale, mahsup...",
        key="manuel_odeme",
    )

    st.caption(
        "Bir alana yazdıktan sonra Enter'a basın veya başka bir "
        "alana tıklayın. Hesaplanan değer otomatik olarak görünecektir."
    )

    # -------------------------------------------------------
    # PDF OLUŞTURMA BUTONU
    # -------------------------------------------------------

    manuel_olustur = st.button(
        "📄 Manuel PDF Oluştur",
        type="primary",
        use_container_width=True,
        key="manuel_pdf_olustur",
    )

    # -------------------------------------------------------
    # FORM KONTROLLERİ
    # -------------------------------------------------------

    if manuel_olustur:

        hatalar = []

        temiz_urun = manuel_urun.strip()
        temiz_isim = manuel_isim.strip()
        temiz_tc = manuel_tc.strip()
        temiz_odeme = manuel_odeme.strip()

        if not temiz_urun:
            hatalar.append(
                "Satılan cinsi boş bırakılamaz."
            )

        if not temiz_isim:
            hatalar.append(
                "İsim soyisim boş bırakılamaz."
            )

        if not temiz_tc:
            hatalar.append(
                "T.C. kimlik numarası boş bırakılamaz."
            )

        elif not temiz_tc.isdigit():
            hatalar.append(
                "T.C. kimlik numarası sadece rakamlardan oluşmalıdır."
            )

        elif len(temiz_tc) != 11:
            hatalar.append(
                "T.C. kimlik numarası 11 rakamdan oluşmalıdır."
            )

        if manuel_miktar <= 0:
            hatalar.append(
                "Miktar sıfırdan büyük olmalıdır."
            )

        if manuel_birim_fiyat <= 0:
            hatalar.append(
                "Birim fiyat sıfırdan büyük olmalıdır."
            )

        if manuel_toplam <= 0:
            hatalar.append(
                "Toplam tutar sıfırdan büyük olmalıdır."
            )

        if not temiz_odeme:
            hatalar.append(
                "Ödeme şekli boş bırakılamaz."
            )

        # ---------------------------------------------------
        # HATALARI GÖSTER
        # ---------------------------------------------------

        if hatalar:

            st.error(
                "Lütfen aşağıdaki alanları kontrol edin:"
            )

            for hata in hatalar:
                st.warning(hata)

        # ---------------------------------------------------
        # PDF'Yİ OLUŞTUR
        # ---------------------------------------------------

        else:

            try:

                manuel_pdf = manuel_pdf_olustur(
                    tarih=manuel_tarih,
                    isim=temiz_isim,
                    tc=temiz_tc,
                    urun=temiz_urun,
                    miktar=manuel_miktar,
                    birim_fiyat=manuel_birim_fiyat,
                    toplam_tutar=manuel_toplam,
                    odeme_sekli=temiz_odeme,
                )

                st.success(
                    "Manuel gider pusulası başarıyla hazırlandı."
                )

                st.download_button(
                    label="⬇️ Manuel Gider Pusulasını İndir",
                    data=manuel_pdf,
                    file_name="Manuel_Gider_Pusulasi.pdf",
                    mime="application/pdf",
                    type="primary",
                    use_container_width=True,
                    key="manuel_pdf_indir",
                )

            except Exception as hata:

                st.error(
                    f"Manuel PDF oluşturulamadı: {hata}"
                )

# ===========================================================
# YAZDIRMA UYARISI
# ===========================================================

st.divider()

st.warning(
    "PDF'yi yazdırırken ölçek ayarını Gerçek Boyut veya %100 seçin. "
    "Sayfaya Sığdır seçeneğini kapatın."
)
