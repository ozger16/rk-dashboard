import streamlit as st
import requests
import pandas as pd
import time
import datetime as dt
from datetime import datetime, timedelta, timezone
import os
from dotenv import load_dotenv
from streamlit_autorefresh import st_autorefresh

# ==========================================
# 0. .ENV DOSYASINDAN BİLGİLERİ ÇEKME
# ==========================================
load_dotenv()

SUPPLIER_ID = os.getenv("SUPPLIER_ID")
API_KEY = os.getenv("API_KEY")
API_SECRET = os.getenv("API_SECRET")
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

SABIT_KARGO_UCRETI = 100.00
SABIT_PLATFORM_BEDELI = 13.19

# Türkiye Saat Dilimini Sabitliyoruz
TR_TZ = timezone(timedelta(hours=3))

st.set_page_config(page_title="Canlı Performans & Kâr Paneli", layout="wide")

if not SUPPLIER_ID or not API_KEY or not API_SECRET:
    st.error("⚠️ API bilgileri `.env` dosyasından okunamadı! Lütfen `.env` dosyanızı kontrol edin.")
    st.stop()

# ==========================================
# 1. YARDIMCI FONKSİYONLAR
# ==========================================
@st.cache_data(ttl=60)
def maliyetleri_excelden_al(dosya_yolu="maliyetler.xlsx"):
    try:
        df = pd.read_excel(dosya_yolu)
        urun_maliyetleri = {}
        for index, row in df.iterrows():
            barkod = str(row['Barkod']).strip()
            urun_maliyetleri[barkod] = {
                "maliyet": float(row['Maliyet']),
                "komisyon_orani": float(row['Komisyon_Orani'])
            }
        return urun_maliyetleri
    except Exception as e:
        return {}

def get_timestamps(periyot, baslangic_tarihi=None, bitis_tarihi=None):
    su_an = datetime.now(TR_TZ)
    
    if periyot == "Bugün (Canlı)":
        baslangic = su_an.replace(hour=0, minute=0, second=0, microsecond=0)
        bitis = su_an + timedelta(minutes=5)
    elif periyot == "Dün":
        dun = su_an - timedelta(days=1)
        baslangic = dun.replace(hour=0, minute=0, second=0, microsecond=0)
        bitis = dun.replace(hour=23, minute=59, second=59, microsecond=999999)
    elif periyot == "Bu Hafta":
        baslangic = (su_an - timedelta(days=6)).replace(hour=0, minute=0, second=0, microsecond=0)
        bitis = su_an + timedelta(minutes=5)
    elif periyot == "Bu Ay":
        baslangic = (su_an - timedelta(days=29)).replace(hour=0, minute=0, second=0, microsecond=0)
        bitis = su_an + timedelta(minutes=5)
    elif periyot == "Özel Tarih" and baslangic_tarihi and bitis_tarihi:
        baslangic = datetime.combine(baslangic_tarihi, dt.time.min).replace(tzinfo=TR_TZ)
        bitis = datetime.combine(bitis_tarihi, dt.time.max).replace(tzinfo=TR_TZ)
    else:
        baslangic = su_an.replace(hour=0, minute=0, second=0, microsecond=0)
        bitis = su_an + timedelta(minutes=5)
        
    tarih_metni = f"{baslangic.strftime('%d.%m.%Y %H:%M:%S')} - {bitis.strftime('%d.%m.%Y %H:%M:%S')}"
    return baslangic, bitis, tarih_metni

def siparisleri_getir(baslangic_dt, bitis_dt):
    url = f"https://api.trendyol.com/sapigw/suppliers/{SUPPLIER_ID}/orders"
    tum_siparisler = []
    
    mevcut_bas = baslangic_dt
    while mevcut_bas < bitis_dt:
        mevcut_bit = mevcut_bas + timedelta(days=14)
        if mevcut_bit > bitis_dt:
            mevcut_bit = bitis_dt
            
        start_ms = int(mevcut_bas.timestamp() * 1000)
        end_ms = int(mevcut_bit.timestamp() * 1000)
        page = 0
        
        while True:
            params = {
                "startDate": start_ms, 
                "endDate": end_ms, 
                "size": 100,
                "page": page,
                "orderByField": "CreatedDate",
                "orderByDirection": "DESC"
            }
            response = requests.get(url, params=params, auth=(API_KEY, API_SECRET))
            
            if response.status_code == 200:
                data = response.json()
                icerik = data.get("content", [])
                tum_siparisler.extend(icerik)
                
                if len(icerik) < 100:
                    break
                page += 1
                time.sleep(0.05)
            else:
                break
                
        mevcut_bas = mevcut_bit
        time.sleep(0.05)
        
    return tum_siparisler

def telegram_mesaj_gonder(mesaj):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        return False, "Telegram Token veya Chat ID .env dosyasında eksik!"
    
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": mesaj,
        "parse_mode": "Markdown"
    }
    try:
        res = requests.post(url, json=payload)
        if res.status_code == 200:
            return True, "Rapor başarıyla gönderildi!"
        else:
            return False, f"Telegram API Hatası: {res.text}"
    except Exception as e:
        return False, f"Bağlantı Hatası: {e}"

def performansi_hesapla(baslangic_dt, bitis_dt, periyot, urun_maliyetleri):
    siparisler = siparisleri_getir(baslangic_dt, bitis_dt)
    
    ozet = {
        "brut_ciro": 0.0, "net_siparis_tutari": 0.0, "net_kar": 0.0, 
        "satilan_urun_adedi": 0, "net_siparis_adedi": 0
    }
    
    tablo_verileri = []
    islenen_paketler = set()
    
    for siparis in siparisler:
        s_id = siparis.get("id")
        if s_id in islenen_paketler:
            continue
        islenen_paketler.add(s_id)

        siparis_tarih_ms = siparis.get("orderDate")
        
        if siparis_tarih_ms:
            utc_dt = datetime.fromtimestamp(siparis_tarih_ms / 1000.0, tz=timezone.utc)
            siparis_dt = datetime(utc_dt.year, utc_dt.month, utc_dt.day, 
                                  utc_dt.hour, utc_dt.minute, utc_dt.second, 
                                  utc_dt.microsecond, tzinfo=TR_TZ)
            
            if periyot == "Bugün (Canlı)":
                if siparis_dt < baslangic_dt:
                    continue
            else:
                if not (baslangic_dt <= siparis_dt <= bitis_dt):
                    continue
                
        status = siparis.get("status")
        if status in ["Cancelled", "UnSupplied"]:
            continue
            
        siparis_no = siparis.get("orderNumber")
        siparis_tarih_str = siparis_dt.strftime('%d.%m.%Y %H:%M:%S') if siparis_tarih_ms else "Bilinmiyor"
        
        siparis_satis_tutari = 0.0
        siparis_urun_maliyeti = 0.0
        siparis_komisyon_tutari = 0.0
        siparis_icindeki_urun_adeti = 0
        gecerli_siparis = False
        
        for urun in siparis.get("lines", []):
            line_status = str(urun.get("orderLineStatusName", "")).lower()
            if "iade" in line_status or "return" in line_status:
                continue

            barkod = str(urun.get("barcode")).strip()
            satis_fiyati = float(urun.get("price", 0))
            adet = int(urun.get("quantity", 1))
            
            satir_tutari = satis_fiyati * adet
            siparis_icindeki_urun_adeti += adet
            siparis_satis_tutari += satir_tutari
            
            if barkod in urun_maliyetleri:
                veriler = urun_maliyetleri[barkod]
                siparis_urun_maliyeti += (veriler["maliyet"] * adet)
                siparis_komisyon_tutari += (satir_tutari * veriler["komisyon_orani"])
            
            gecerli_siparis = True

        if gecerli_siparis and siparis_icindeki_urun_adeti > 0:
            ozet["net_siparis_adedi"] += 1
            ozet["satilan_urun_adedi"] += siparis_icindeki_urun_adeti

            kargo_kesintisi = SABIT_KARGO_UCRETI
            platform_bedeli = siparis.get("platformServiceFee", SABIT_PLATFORM_BEDELI)
            
            net_siparis_tutari = siparis_satis_tutari - siparis_komisyon_tutari - kargo_kesintisi - platform_bedeli
            siparis_kari = net_siparis_tutari - siparis_urun_maliyeti
            
            ozet["brut_ciro"] += siparis_satis_tutari
            ozet["net_siparis_tutari"] += net_siparis_tutari
            ozet["net_kar"] += siparis_kari
            
            tablo_verileri.append({
                "Sipariş Tarihi": siparis_tarih_str,
                "Sipariş No": siparis_no,
                "Adet": siparis_icindeki_urun_adeti,
                "Brüt Satış (TL)": round(siparis_satis_tutari, 2),
                "Maliyet (TL)": round(siparis_urun_maliyeti, 2),
                "Komisyon (TL)": round(siparis_komisyon_tutari, 2),
                "Kargo & Hizmet (TL)": round(kargo_kesintisi + platform_bedeli, 2),
                "Hesaba Yatacak (TL)": round(net_siparis_tutari, 2),
                "Net KÂR (TL)": round(siparis_kari, 2)
            })

    return ozet, tablo_verileri

# ==========================================
# 2. ANA UYGULAMA VE SEKMELER
# ==========================================

URUN_MALIYETLERI = maliyetleri_excelden_al("maliyetler.xlsx")

sekme1, sekme2 = st.tabs(["📊 Canlı Performans Paneli", "💡 Trendyol Fiyat & Komisyon Simülatörü"])

with sekme1:
    col_baslik, col_buton = st.columns([4, 1])
    with col_baslik:
        st.title("🔥 Canlı Performansım")
    with col_buton:
        st.write("")
        if st.button("🔄 Verileri Yenile", use_container_width=True, type="primary"):
            st.cache_data.clear()

    with st.sidebar:
        st.header("⚙️ Tarih Filtresi")
        secilen_periyot = st.selectbox("Zaman Dilimi Seçin", 
                                       ["Bugün (Canlı)", "Dün", "Bu Hafta", "Bu Ay", "Özel Tarih"])
        
        if secilen_periyot == "Bugün (Canlı)":
            st_autorefresh(interval=10000, key="canli_otomatik_yenileme")
            st.caption("🟢 Canlı mod aktif: Her 10 saniyede bir güncelleniyor.")
        
        baslangic_tarihi, bitis_tarihi = None, None
        if secilen_periyot == "Özel Tarih":
            baslangic_tarihi = st.date_input("Başlangıç Tarihi")
            bitis_tarihi = st.date_input("Bitiş Tarihi")
        
        st.markdown("---")
        
        st.subheader("📱 Bildirimler")
        if st.button("📤 Telegram'a Özet Gönder", use_container_width=True):
            b_dt, bit_dt, t_metin = get_timestamps(secilen_periyot, baslangic_tarihi, bitis_tarihi)
            ozt, _ = performansi_hesapla(b_dt, bit_dt, secilen_periyot, URUN_MALIYETLERI)
            
            rapor_metni = (
                f"📊 *Trendyol {secilen_periyot} Raporu*\n\n"
                f"📦 Sipariş Adedi: *{ozt['net_siparis_adedi']}*\n"
                f"🛍️ Satılan Ürün: *{ozt['satilan_urun_adedi']}*\n"
                f"💳 Brüt Ciro: *{ozt['brut_ciro']:,.2f} ₺*\n"
                f"🏦 Hesaba Yatacak: *{ozt['net_siparis_tutari']:,.2f} ₺*\n"
                f"💰 NET KÂR: *{ozt['net_kar']:,.2f} ₺*\n\n"
                f"🕒 Zaman: {datetime.now(TR_TZ).strftime('%d.%m.%Y %H:%M:%S')}"
            )
            basarili, mesaj = telegram_mesaj_gonder(rapor_metni)
            if basarili:
                st.success(mesaj)
            else:
                st.error(mesaj)

        st.markdown("---")
        if URUN_MALIYETLERI:
            st.success(f"✅ Excel'den {len(URUN_MALIYETLERI)} ürün aktif.")

    if URUN_MALIYETLERI:
        try:
            baslangic_dt, bitis_dt, tarih_bilgisi = get_timestamps(secilen_periyot, baslangic_tarihi, bitis_tarihi)
            
            st.sidebar.info(f"**Taranan Zaman Aralığı:**\n\n{tarih_bilgisi}")
            
            ozet, tablo = performansi_hesapla(baslangic_dt, bitis_dt, secilen_periyot, URUN_MALIYETLERI)
            
            st.markdown("### 📊 Genel Özet")
            kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
            
            kpi1.metric("📦 Sipariş Adedi", f"{ozet['net_siparis_adedi']}")
            kpi2.metric("🛍️ Satılan Ürün", f"{ozet['satilan_urun_adedi']}")
            kpi3.metric("💳 Brüt Ciro", f"{ozet['brut_ciro']:,.2f} ₺", help="Müşterinin ödediği toplam tutar")
            kpi4.metric("🏦 Hesaba Yatacak", f"{ozet['net_siparis_tutari']:,.2f} ₺", help="Trendyol kesintileri sonrası bankanıza gelecek para")
            kpi5.metric("💰 NET KÂR", f"{ozet['net_kar']:,.2f} ₺", help="Hesaba yatacak tutar eksi ürün alış maliyetiniz")
            
            st.markdown("---")
            
            st.markdown("#### 📋 Son Siparişler ve Kârlılık Detayı")
            if tablo:
                df_tablo = pd.DataFrame(tablo)
                st.dataframe(df_tablo, use_container_width=True, hide_index=True)
            else:
                st.info("Bu periyodda gösterilecek sipariş yok.")
                    
        except Exception as e:
            st.error(f"Veriler hesaplanırken bir hata oluştu: {e}")
    else:
        st.warning("Lütfen maliyetler.xlsx dosyanızı hazırlayın.")

# ==========================================
# 3. TRENDYOL KOMİSYON EXCELİ DESTEKLİ SİMÜLATÖR
# ==========================================
with sekme2:
    st.title("💡 Trendyol Fiyat & Komisyon Simülatörü")
    st.markdown("Trendyol'dan indirdiğiniz **Komisyon Tarifeleri** Excel dosyasını yükleyerek tüm fiyat aralıklarını ve net kârınızı otomatik simüle edin.")
    
    trendyol_excel_dosyasi = st.file_uploader(
        "Trendyol Komisyon Tarifeleri Excel Dosyasını Yükle (.xlsx)", 
        type=["xlsx"],
        key="ty_excel"
    )
    
    if trendyol_excel_dosyasi is not None:
        try:
            ty_df = pd.read_excel(trendyol_excel_dosyasi)
            st.success("✅ Trendyol komisyon raporu başarıyla yüklendi!")
            
            # Ürün seçimi
            barkod_listesi = ty_df['BARKOD'].astype(str).tolist()
            urun_isimleri = ty_df['ÜRÜN İSMİ'].astype(str).tolist()
            secenekler = [f"{b} - {i}" for b, i in zip(barkod_listesi, urun_isimleri)]
            
            secilen_secenek = st.selectbox("Simülasyon Yapılacak Ürünü Seçin", secenekler)
            
            if secilen_secenek:
                secilen_barkod = secilen_secenek.split(" - ")[0]
                urun_satiri = ty_df[ty_df['BARKOD'].astype(str) == secilen_barkod].iloc[0]
                
                # Maliyeti bul
                maliyet = 0.0
                if secilen_barkod in URUN_MALIYETLERI:
                    maliyet = URUN_MALIYETLERI[secilen_barkod]["maliyet"]
                
                # Trendyol verilerini çek
                guncel_tsf = float(urun_satiri.get('GÜNCEL TSF', 0))
                guncel_kom = float(urun_satiri.get('GÜNCEL KOMİSYON', 0))
                
                k1_fiyat = float(urun_satiri.get('KOMİSYONA ESAS FİYAT', guncel_tsf))
                k1_kom = float(urun_satiri.get('1.KOMİSYON', guncel_kom)) / 100.0
                
                k2_fiyat = float(urun_satiri.get('2.Fiyat Alt Limit', guncel_tsf * 0.9))
                k2_kom = float(urun_satiri.get('2.KOMİSYON', guncel_kom - 1)) / 100.0
                
                k3_fiyat = float(urun_satiri.get('3.Fiyat Alt Limit', guncel_tsf * 0.8))
                k3_kom = float(urun_satiri.get('3.KOMİSYON', guncel_kom - 2)) / 100.0
                
                k4_fiyat = float(urun_satiri.get('4.Fiyat Üst Limiti', guncel_tsf * 0.7))
                k4_kom = float(urun_satiri.get('4.KOMİSYON', guncel_kom - 3)) / 100.0
                
                st.markdown("---")
                st.info(f"📦 **Ürün:** {urun_satiri.get('ÜRÜN İSMİ')} | **Mevcut Alış Maliyeti:** **{maliyet:,.2f} ₺** (Eğer maliyet girilmediyse 0 alınır)")
                
                st.subheader("📊 Trendyol Fiyat Aralıkları ve Kâr Simülasyonu")
                
                def hesapla(fiyat, kom_orani):
                    kom_tutar = fiyat * kom_orani
                    hizmet = SABIT_KARGO_UCRETI + SABIT_PLATFORM_BEDELI
                    yatacak = fiyat - kom_tutar - hizmet
                    net_k = yatacak - maliyet
                    marj = (net_k / fiyat * 100) if fiyat > 0 else 0
                    return kom_tutar, yatacak, net_k, marj

                s1_kom, s1_yat, s1_kar, s1_marj = hesapla(k1_fiyat, k1_kom)
                s2_kom, s2_yat, s2_kar, s2_marj = hesapla(k2_fiyat, k2_kom)
                s3_kom, s3_yat, s3_kar, s3_marj = hesapla(k3_fiyat, k3_kom)
                s4_kom, s4_yat, s4_kar, s4_marj = hesapla(k4_fiyat, k4_kom)
                
                sim_tablo = pd.DataFrame({
                    "Aralık / Senaryo": ["1. Fiyat Aralığı", "2. Fiyat Aralığı", "3. Fiyat Aralığı", "4. Fiyat Aralığı"],
                    "Satış Fiyatı (TL)": [k1_fiyat, k2_fiyat, k3_fiyat, k4_fiyat],
                    "Komisyon Oranı": [f"%{k1_kom*100:.1f}", f"%{k2_kom*100:.1f}", f"%{k3_kom*100:.1f}", f"%{k4_kom*100:.1f}"],
                    "Komisyon Kesintisi (TL)": [round(s1_kom, 2), round(s2_kom, 2), round(s3_kom, 2), round(s4_kom, 2)],
                    "Hesaba Yatacak (TL)": [round(s1_yat, 2), round(s2_yat, 2), round(s3_yat, 2), round(s4_yat, 2)],
                    "Net Kâr (TL)": [round(s1_kar, 2), round(s2_kar, 2), round(s3_kar, 2), round(s4_kar, 2)],
                    "Kâr Marjı": [f"%{s1_marj:.1f}", f"%{s2_marj:.1f}", f"%{s3_marj:.1f}", f"%{s4_marj:.1f}"]
                })
                
                st.dataframe(sim_tablo, use_container_width=True, hide_index=True)
                
                en_iyi_kar = max(s1_kar, s2_kar, s3_kar, s4_kar)
                st.success(f"🎯 Bu ürün için en yüksek net kâr (**{en_iyi_kar:,.2f} ₺**) yukarıdaki aralıklardan en karlı olan fiyatta elde edilir.")
                
        except Exception as e:
            st.error(f"Excel dosyası okunurken hata oluştu: {e}")
    else:
        st.info("💡 Başlamak için lütfen yukarıdan Trendyol'dan indirdiğiniz komisyon tarifeleri Excel dosyanızı yükleyin.")
