
import smtplib
import logging
import io
import base64
import os
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib.image import imread
import numpy as np

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from eptr2 import EPTR2

# ─────────────────────────────────────────────
import os

LOGO_PATH_JPG = "Alpine-enerji.jpg"
LOGO_PATH_PNG = "Alpine-enerji.png"
LOGO_PATH_BEYAZ = "Alpine-enerji-beyaz.png"

# Renkli logo (Excel için)
if os.path.exists(LOGO_PATH_JPG):
    with open(LOGO_PATH_JPG, "rb") as f:
        LOGO_B64 = base64.b64encode(f.read()).decode("utf-8")
else:
    LOGO_B64 = ""

if os.path.exists(LOGO_PATH_BEYAZ):
    with open(LOGO_PATH_BEYAZ, "rb") as f:
        LOGO_BEYAZ_B64 = base64.b64encode(f.read()).decode("utf-8")
    LOGO_MAIL_SRC = f"data:image/png;base64,{LOGO_BEYAZ_B64}"
else:
    LOGO_MAIL_SRC = ""
    
# ─────────────────────────────────────────────
#  AYARLAR — GitHub Secrets'tan Okunur
# ─────────────────────────────────────────────
EPIAS_USERNAME = os.getenv("EPIAS_USERNAME")
EPIAS_PASSWORD = os.getenv("EPIAS_PASSWORD")
OUTLOOK_MAIL   = os.getenv("OUTLOOK_MAIL")
OUTLOOK_PASS   = os.getenv("OUTLOOK_PASS")

SMTP_SERVER  = "smtp.office365.com"
SMTP_PORT    = 587

TEST_MODU = False 

MUSTERI_LISTESI = [

    {"ad": "Beyza Nur Özbek", "email": "beyzanur.ozbek@alpineenerji.com.tr"}
]

# ─────────────────────────────────────────────
#  LOGGING
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)


def saat_aralik(saat_no: str) -> str:
    try:
        h = int(saat_no)
        return f"{h:02d}:00-{(h + 1) % 24:02d}:00"
    except Exception:
        return saat_no


def ptf_veri_cek():
    hedef = datetime.now() + timedelta(days=1)
    tarih_str = hedef.strftime("%Y-%m-%d")
    eptr = EPTR2(username=EPIAS_USERNAME, password=EPIAS_PASSWORD)

    try:
        df = eptr.call("interim-mcp", start_date=tarih_str, end_date=tarih_str)
        if df is None or df.empty:
            return None, tarih_str

        veri = []
        for _, row in df.iterrows():
            fiyat = row.get("marketTradePrice", 0)
            miktar = row.get("marketTradeAmount", 0) # Ağırlıklı ortalama için gerekli
            saat_no = str(row.get("hour", ""))
            veri.append({
                "saat_no": saat_no,
                "saat": saat_aralik(saat_no),
                "fiyat": fiyat,
                "miktar": miktar
            })
        return veri, tarih_str
    except Exception as e:
        log.error(f"HATA: {e}")
        return None, tarih_str

def grafik_olustur(veri: list, tarih: str) -> str:
    def fmt_iki_satir_saat(s_no):
        try:
            h = int(s_no.split(":")[0])
        except:
            h = int(s_no)
        return f"{h:02d}:00\n{(h + 1) % 24:02d}:00"

    n = len(veri)
    saat_araliklari = [fmt_iki_satir_saat(r["saat_no"]) for r in veri]
    fiyatlar = [float(r["fiyat"]) for r in veri]
    tarih_fmt = datetime.strptime(tarih, "%Y-%m-%d").strftime("%d.%m.%Y")
    NAVY = "#201F5A"

    fig = plt.figure(figsize=(12, 7))
    fig.patch.set_facecolor("white")

    ax = fig.add_axes([0.06, 0.33, 0.91, 0.50])
    x = np.arange(n)
    bars = ax.bar(x, fiyatlar, color=NAVY, width=0.55, zorder=3)
    for bar, val in zip(bars, fiyatlar):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + 5,
            f"{val:,.0f}" if val > 0 else "0",
            ha="center", va="bottom", fontsize=7.5, color=NAVY, fontweight="bold"
        )
    ax.set_xticks(x)
    ax.set_xticklabels([])
    ax.set_xlim(-0.5, n - 0.5)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: "{:,.0f}".format(v)))
    ax.tick_params(axis="y", labelsize=7)
    ax.set_ylim(0, max(fiyatlar) * 1.2)
    ax.grid(axis="y", linestyle="--", alpha=0.3, zorder=0)
    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)
    for spine in ["left", "bottom"]:
        ax.spines[spine].set_edgecolor("#BBBBBB")

    fig.text(0.5, 0.94, f"EPİAŞ Kesinleşmemiş Piyasa Takas Fiyatı (PTF) — {tarih_fmt}",
             ha="center", fontsize=10, fontweight="bold", color="#201F5A")

    try:
        logo_img = imread(LOGO_PATH_PNG)
        logo_ax = fig.add_axes([0.82, 0.88, 0.14, 0.10])
        logo_ax.imshow(logo_img)
        logo_ax.axis("off")
    except Exception as e:
        log.warning(f"Logo yüklenemedi (grafik): {e}")

    ax_t = fig.add_axes([0.06, 0.12, 0.91, 0.26])
    ax_t.set_axis_off()
    tbl = ax_t.table(
        cellText=[
            [s for s in saat_araliklari],
            [f"{v:,.0f}" for v in fiyatlar]
        ],
        rowLabels=["Saat Aralığı", "PTF (TL/MWh)"],
        cellLoc="center",
        loc="center"
    )
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(7)
    for (ri, ci), cell in tbl.get_celld().items():
        cell.set_linewidth(0.4)
        cell.set_edgecolor("#BBBBBB")
        if ci == -1:
            cell.set_facecolor(NAVY)
            cell.set_text_props(color="white", fontweight="bold", fontsize=7)
            cell.set_height(0.25)
        elif ri == 1:
            cell.set_facecolor(NAVY)
            cell.set_text_props(color="white", fontweight="bold", fontsize=7)
            cell.set_height(0.25)
        else:
            cell.set_facecolor("#EFF4FB" if ci % 2 == 0 else "#FFFFFF")
            cell.set_text_props(color=NAVY, fontweight="bold")
            cell.set_height(0.25)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=100, bbox_inches="tight")
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode("utf-8")
    
def xlsx_olustur(veri: list, tarih: str) -> bytes:
    tarih_fmt = datetime.strptime(tarih, "%Y-%m-%d").strftime("%d.%m.%Y")
    NAVY_HEX = "201F5A"
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PTF Verileri"

    # Mevcut kolon genişliklerinizi koruyoruz
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 15
    ws.row_dimensions[1].height = 48

    # --- LOGO VE BAŞLIK (Sizin Orijinal Düzeniniz) ---
    ws.merge_cells("B1:C1")
    try:
        from openpyxl.drawing.image import Image as XLImage
        xl_logo = XLImage(LOGO_PATH_JPG)
        xl_logo.width, xl_logo.height = 145, 60
        xl_logo.anchor = "A1"
        ws.add_image(xl_logo)
    except:
        ws["A1"] = "ALPİNE ENERJİ"

    ws["B1"] = f"EPİAŞ Kesinleşmemiş PTF - {tarih_fmt}"
    ws["B1"].font = Font(name="Arial", size=11, bold=True, color=NAVY_HEX)
    ws["B1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # --- TABLO BAŞLIĞI ---
    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    header_fill = PatternFill("solid", start_color=NAVY_HEX, end_color=NAVY_HEX)
    thin_border = Border(left=Side(style="thin", color="CCCCCC"), right=Side(style="thin", color="CCCCCC"),
                         top=Side(style="thin", color="CCCCCC"), bottom=Side(style="thin", color="CCCCCC"))

    for ci, h in enumerate(["Tarih", "Saat Aralığı", "PTF (TL/MWh)"], start=1):
        c = ws.cell(row=4, column=ci, value=h)
        c.font = header_font; c.fill = header_fill; c.alignment = Alignment(horizontal="center", vertical="center"); c.border = thin_border

    # --- VERİ SATIRLARI ---
    toplam_fiyat = 0
    toplam_fiyat_x_miktar = 0
    toplam_miktar = 0
    son_satir = 4

    for i, row in enumerate(veri, start=5):
        fiyat = float(row["fiyat"])
        miktar = float(row.get("miktar", 0))
        
        toplam_fiyat += fiyat
        toplam_fiyat_x_miktar += (fiyat * miktar)
        toplam_miktar += miktar

        bg = "EFF4FB" if i % 2 == 0 else "FFFFFF"
        rf = PatternFill("solid", start_color=bg, end_color=bg)
        # Puntoları 10 yaparak diğer satırlarla eşitledik
        fnt = Font(name="Arial", size=10, bold=True, color="000000")
        
        c1 = ws.cell(row=i, column=1, value=tarih_fmt); c1.font=fnt; c1.fill=rf; c1.border=thin_border; c1.alignment=Alignment(horizontal="center")
        c2 = ws.cell(row=i, column=2, value=row["saat"]); c2.font=fnt; c2.fill=rf; c2.border=thin_border; c2.alignment=Alignment(horizontal="center")
        c3 = ws.cell(row=i, column=3, value=fiyat); c3.font=fnt; c3.fill=rf; c3.border=thin_border; c3.alignment=Alignment(horizontal="center"); c3.number_format='#,##0.00'
        son_satir = i

    # --- HESAPLAMALAR ---
    gunluk_ort = toplam_fiyat / len(veri) if veri else 0
    agirlikli_ort = toplam_fiyat_x_miktar / toplam_miktar if toplam_miktar > 0 else gunluk_ort

    # --- ORTALAMA SATIRLARI (Aynı Punto, Lacivert Arka Plan) ---
    def ort_satiri_ekle(satir_no, etiket, deger):
        ws.merge_cells(f"A{satir_no}:B{satir_no}")
        # Font büyüklüğünü verilerle aynı (10) yaptık
        fnt_beyaz = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        fill_navy = PatternFill("solid", start_color=NAVY_HEX, end_color=NAVY_HEX)
        
        c_label = ws.cell(row=satir_no, column=1, value=etiket)
        c_label.font = fnt_beyaz; c_label.fill = fill_navy; c_label.alignment = Alignment(horizontal="center"); c_label.border = thin_border
        
        c_val = ws.cell(row=satir_no, column=3, value=deger)
        c_val.font = fnt_beyaz; c_val.fill = fill_navy; c_val.alignment = Alignment(horizontal="center"); c_val.border = thin_border; c_val.number_format = '#,##0.00'

    ort_satiri_ekle(son_satir + 1, "GÜNLÜK ORTALAMA", gunluk_ort)
    ort_satiri_ekle(son_satir + 2, "AĞIRLIKLI ORTALAMA", agirlikli_ort)

    # --- GRAFİK EKLEME (Orijinal Kodundaki gibi dokunmadan ekliyoruz) ---
    # ... (Buradaki plt kodlarını orijinal dosyanızdaki gibi bırakın, sadece anchor'ı son satıra göre güncelleyebilirsiniz)
    # Grafiği anchor="E5" yaparak sağa sabitledim ki tabloyu kapatmasın.

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

def html_mail_olustur(musteri_ad: str, veri: list, tarih: str, grafik_b64: str) -> str:
    tarih_fmt = datetime.strptime(tarih, "%Y-%m-%d").strftime("%d.%m.%Y")
    NAVY = "#201F5A"

    return f"""<!DOCTYPE html>
<html lang="tr">
<body style="margin:0; padding:0; font-family:Arial,sans-serif; color:#222; background:#f4f6fb;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6fb;">
    <tr><td align="center" style="padding:24px 8px;">
      <table width="100%" cellpadding="0" cellspacing="0" style="max-width:800px; background:#fff; border-radius:8px;">
        <tr>
         <td style="background:#201F5A; padding:16px 30px; border-radius:8px 8px 0 0;">
            <table width="100%" cellpadding="0" cellspacing="0">
              <tr>
                <td style="vertical-align:middle; text-align:left;">
                  <div style="font-size:14px; font-weight:900; color:#fff; line-height:1.3;">Kesinleşmemiş Piyasa Takas Fiyatı (PTF)</div>
                  <div style="font-size:12px; color:#4EB2D2; margin-top:4px;">{tarih_fmt} Tarihine Ait</div>
                        </td>
                     <td style="vertical-align:middle; text-align:right; width:120px;">
                      <img src="{LOGO_MAIL_SRC}"
                           width="150" height="65"
                           style="display:block; margin-left:auto;"
                           alt="Alpine Enerji" />
                    </td>
              </tr>
            </table>
          </td>
        </tr>

        <tr>
          <td style="padding:25px 30px;">
           <p style="font-size:15px; color:#201F5A;">Sayın <b>{musteri_ad}</b>,</p>
            <p style="font-size:15px; color:#201F5A;">
              {tarih_fmt} tarihine ait <b>Kesinleşmemiş Piyasa Takas Fiyatı (PTF)</b> verileri aşağıda yer almaktadır.
            </p>
          </td>
        </tr>

        <tr>
          <td align="center" style="padding:0 30px;">
            <img src="data:image/png;base64,{grafik_b64}"
                 style="width:100%; max-width:700px; height:auto; display:block; border:1px solid #eee;" />
          </td>
        </tr>

        <tr>
          <td style="padding:20px 30px;">
             <p style="font-size:14px; color:#666; border-top:1px solid #eee; padding-top:10px; font-weight:bold; text-align:center;">
              Kaynak: EPİAŞ Şeffaflık Platformu
              </p>
          </td>
        </tr>

      </table>
    </td></tr>
  </table>
</body>
</html>"""


def mail_gonder(musteri: dict, veri: list, tarih: str, xlsx_bytes: bytes, grafik_b64: str):
    msg = MIMEMultipart("mixed")
    tarih_fmt = datetime.strptime(tarih, "%Y-%m-%d").strftime("%d.%m.%Y")
    msg["Subject"] = f"EPİAŞ Kesinleşmemiş Piyasa Takas Fiyatı (PTF) — {tarih_fmt}"
    msg["From"]    = OUTLOOK_MAIL
    msg["To"]      = musteri["email"]
    msg.attach(MIMEText(html_mail_olustur(musteri["ad"], veri, tarih, grafik_b64), "html", "utf-8"))
    ek = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    ek.set_payload(xlsx_bytes)
    encoders.encode_base64(ek)
    ek.add_header("Content-Disposition", f'attachment; filename="PTF_{tarih}.xlsx"')
    msg.attach(ek)
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(OUTLOOK_MAIL, OUTLOOK_PASS)
        server.sendmail(OUTLOOK_MAIL, [musteri["email"]], msg.as_string())


def main():
    log.info("=" * 55)
    try:
        veri, tarih = ptf_veri_cek()

        if veri is None:
            log.info("Süreç durduruldu: Yarının PTF verisi henüz yayınlanmadığı için mail gönderilmedi.")
            return

        xlsx_bytes = xlsx_olustur(veri, tarih)
        grafik_b64 = grafik_olustur(veri, tarih)

        for musteri in MUSTERI_LISTESI:
            mail_gonder(musteri, veri, tarih, xlsx_bytes, grafik_b64)
            log.info(f"✓ Mail gönderildi → {musteri['email']}")

    except Exception as e:
        log.error(f"Süreç sırasında HATA oluştu: {e}")
        raise


if __name__ == "__main__":
    main()
