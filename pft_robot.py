import smtplib
import logging
import io
import base64
import os
import sys
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from eptr2 import EPTR2

# ─────────────────────────────────────────────
#  AYARLAR & GITHUB SECRETS
# ─────────────────────────────────────────────
EPIAS_USERNAME = os.getenv("EPIAS_USERNAME")
EPIAS_PASSWORD = os.getenv("EPIAS_PASSWORD")
OUTLOOK_MAIL   = os.getenv("OUTLOOK_MAIL")
OUTLOOK_PASS   = os.getenv("OUTLOOK_PASS")
SMTP_SERVER    = "smtp.office365.com"
SMTP_PORT      = 587

# --- MÜŞTERİ LİSTESİ (Hata burada giderildi) ---
MUSTERILER = [
    {"ad": "Beyzanur Özbek", "email": "beyzanur.ozbek@alpineenerji.com.tr"}
]

# ─────────────────────────────────────────────
#  LOGGING & YARDIMCI FONKSİYONLAR
# ─────────────────────────────────────────────
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

def saat_aralik(saat_no: str) -> str:
    try:
        h = int(saat_no)
        return f"{h:02d}:00-{(h + 1) % 24:02d}:00"
    except:
        return saat_no

# ─────────────────────────────────────────────
#  VERİ ÇEKME, GRAFİK VE EXCEL (Sistem Fonksiyonları)
# ─────────────────────────────────────────────
def pft_veri_cek():
    hedef = datetime.now() + timedelta(days=1)
    tarih_str = hedef.strftime("%Y-%m-%d")
    log.info(f"Yarının verisi çekiliyor: {tarih_str}")
    
    eptr = EPTR2(username=EPIAS_USERNAME, password=EPIAS_PASSWORD)
    df = eptr.call("interim-mcp", start_date=tarih_str, end_date=tarih_str)
    
    veri = []
    for _, row in df.iterrows():
        fiyat = row.get("marketTradePrice", 0.0)
        veri.append({
            "saat_no": str(row.get("hour", "")),
            "saat": saat_aralik(str(row.get("hour", ""))),
            "fiyat": fiyat if fiyat is not None else 0.0
        })
    if not veri:
        raise ValueError(f"{tarih_str} verisi henüz yayınlanmamış.")
    return veri, tarih_str

def grafik_olustur(veri: list, tarih: str) -> str:
    fiyatlar = [float(r["fiyat"]) for r in veri]
    saatler = [r["saat"] for r in veri]
    fig, ax = plt.subplots(figsize=(14, 6))
    ax.plot(range(len(saatler)), fiyatlar, color="#201F5A", linewidth=2, marker='o')
    ax.set_xticks(range(len(saatler)))
    ax.set_xticklabels(saatler, rotation=45, ha="right")
    plt.title(f"PFT - {tarih}")
    plt.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=100)
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode("utf-8")

def xlsx_olustur(veri: list, tarih: str) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Tarih", "Saat", "Fiyat (TL)"])
    for r in veri:
        ws.append([tarih, r["saat"], r["fiyat"]])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ─────────────────────────────────────────────
#  MAIL GÖNDERME
# ─────────────────────────────────────────────
def mail_gonder(musteri: dict, veri: list, tarih: str, xlsx_bytes: bytes, grafik_b64: str):
    msg = MIMEMultipart("mixed")
    msg["Subject"] = f"EPİAŞ Kesinleşmemiş PFT — {tarih}"
    msg["From"] = OUTLOOK_MAIL
    msg["To"] = musteri["email"]

    # HTML İçerik (Kısa versiyon)
    html = f"<html><body><h3>Sayın {musteri['ad']},</h3><p>{tarih} verileri ektedir.</p><img src='data:image/png;base64,{grafik_b64}' width='1000'></body></html>"
    msg.attach(MIMEText(html, "html", "utf-8"))

    # Excel Eki
    ek = MIMEBase("application", "octet-stream")
    ek.set_payload(xlsx_bytes)
    encoders.encode_base64(ek)
    ek.add_header("Content-Disposition", f'attachment; filename="PFT_{tarih}.xlsx"')
    msg.attach(ek)

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(OUTLOOK_MAIL, OUTLOOK_PASS)
        server.sendmail(OUTLOOK_MAIL, musteri["email"], msg.as_string())

# ─────────────────────────────────────────────
#  ANA AKIŞ
# ─────────────────────────────────────────────
def main():
    try:
        veri, tarih = pft_veri_cek()
        g_b64 = grafik_olustur(veri, tarih)
        xlsx = xlsx_olustur(veri, tarih)
        
        for m in MUSTERILER:
            mail_gonder(m, veri, tarih, xlsx, g_b64)
            log.info(f"✓ {m['email']} adresine başarıyla gönderildi.")
            
    except Exception as e:
        log.error(f"Sistem Hatası: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
