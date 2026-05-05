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
#  LOGO — Kod ile aynı klasörde olmalı
# ─────────────────────────────────────────────
LOGO_PATH_JPG = "assets/Alpine-enerji.jpg"
LOGO_PATH_PNG = "assets/Alpine-enerji.png"

# Logo dosyasını base64'e çevir (e-posta HTML için)
with open(LOGO_PATH_JPG, "rb") as f:
    LOGO_B64 = base64.b64encode(f.read()).decode("utf-8")

# E-posta header için beyaz şeffaf logo (gömülü)
LOGO_MAIL_SRC = "data:image/png;base64,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wenTp322A/MmM2bMkP7zn//Q448/rnkWkjO9Sa9evdwoJePLOGxAxOa5nrTIwj3QHctXguXLl0unTp2iKlWqaG4s4tru3bu73ICIkbjZbFY3pwsjwjieO3eOZwYGYtq0aRg9erTubA1du3ZF9erV6fjx41ynzL9wuJeNiYkBoG/z3M/PD7t378a2bdvcqoTiTAStjUW4L1arVg0dO3bMc5qQnZ0NQPsavyiDxx9/XNN9vkixYsW8LQKTgzFjxkhXr15VB3GOIhxMihcvznEhTJ44ZED69etHDz/8sFNr27NmzdJ1nxbmzZune4PfkZiQa9eu6UrvIRpu06ZNNcvlaxTmZTpfZfbs2U6lN+nbt697BMsDMVDTgrcOiyrqOGRAXn75ZV0PF/sl165dQ3JystunwPv27ZPS0tKciglp3759ntekp6dLN2/eBKCtoxRLZD179tQkE8O4gsmTJyM7O1tXkkVFUVChQgW8++67HhsZ3Lhxw+FrJUkCEaF06dKoX78+j148TIEGpH79+hQeHq578xwAli5dqk86HYggRb0xIaVLl8aQIUPyvPnixYuany8CuqpXr45PPvmElZzxKBkZGdLixYt1DazE0pcns/ReunRJ0/ViZaRx48ZukojJiwINSJ8+fdQIcs0P/3vkrefUQb189dVX0rVr15w6J6R37955/i0jIwOAdgMlyvCdd95BYmIiGxHGo3z22We60g+JWUiVKlXQr18/j+htZmYmAMcPqhNLyvm1W8Y9FFhDL774IgDtm+dCWfft24dNmzZ51IPj559/BqA9waLwf3/mmWcQGhp638aye/duAPrW+sXz//Of/2Dnzp0UFxfHhoTxCGlpadL69etdmt7EXYF5GzdulG7evOnwxr9oV+Hh4ejatSu3KQ+S785Tr169qGLFiro2poUBcWXeK0eZM2cOYmNjdSm48H+Pi4u7b+To5s2bMWLECN2NRzTgZ599FrNnz8bnn39O6enpOHz4MM6fP49r167pjqZ3FPGNe/fuxdq1a9k9s4iQlJSENm3aaNYt0UGHhIQgKiqKli1bJgF3D+tyF4cOHcKzzz6rKVGqoihITk52a7wZo4H169cTEZHdbictKIpCREQ3btzw2mggIyODiIhkWdYku7j+3Llzecp+4cKFXN+pB1mWNZerq5k7d+59v/Hs2bNEpO37bDYbEREl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"



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
    hedef = datetime.now()
    tarih_str = hedef.strftime("%Y-%m-%d")
    log.info(f"Bugünün ({tarih_str}) PTF verisi deneniyor... [TEST MODU]")

    eptr = EPTR2(username=EPIAS_USERNAME, password=EPIAS_PASSWORD)

    try:
        df = eptr.call("interim-mcp", start_date=tarih_str, end_date=tarih_str)

        if df is None or df.empty:
            log.warning(f"⚠️ {tarih_str} tarihli PTF verisi henüz EPİAŞ tarafından yayınlanmamış.")
            return None, tarih_str

        veri = []
        for _, row in df.iterrows():
            saat_no = str(row.get("hour", ""))
            fiyat = row.get("marketTradePrice", None)
            veri.append({
                "saat_no": saat_no,
                "saat": saat_aralik(saat_no),
                "fiyat": fiyat if fiyat is not None else 0.0,
                "fiyat_str": str(fiyat) if fiyat is not None else "-",
            })

        if not veri:
            log.warning(f"⚠️ {tarih_str} tarihli PTF verisi boş geldi.")
            return None, tarih_str

        log.info(f"✓ Yarının ({tarih_str}) verisi başarıyla alındı. ({len(veri)} saatlik)")
        return veri, tarih_str

    except Exception as e:
        log.error(f"HATA: Veri çekilirken bir sorun oluştu: {e}")
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

    fig = plt.figure(figsize=(12, 5))
    fig.patch.set_facecolor("white")

    # ── GRAFİK ALANI ──
    ax = fig.add_axes([0.06, 0.24, 0.91, 0.58])
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

    # Başlık
    fig.text(0.5, 0.94, f"EPİAŞ Kesinleşmemiş Piyasa Takas Fiyatı (PTF) — {tarih_fmt}",
             ha="center", fontsize=10, fontweight="bold", color="#222")

    # ── LOGO (sağ üst köşe) ──
    try:
        logo_img = imread(LOGO_PATH_PNG)
        logo_ax = fig.add_axes([0.82, 0.88, 0.14, 0.10])
        logo_ax.imshow(logo_img)
        logo_ax.axis("off")
    except Exception as e:
        log.warning(f"Logo yüklenemedi (grafik): {e}")
        fig.text(0.97, 0.94, "ALPİNE", fontsize=8, fontweight="black", color="#2b2982", ha="right")
        fig.text(0.97, 0.90, "ENERJİ", fontsize=8, fontweight="black", color="#2b2982", ha="right")

    # ── TABLO ALANI ──
    ax_t = fig.add_axes([0.06, 0.04, 0.91, 0.15])
    ax_t.set_axis_off()

    tbl = ax_t.table(
        cellText=[[f"{v:,.2f}" for v in fiyatlar]],
        rowLabels=["PTF (TL/MWh)"],
        colLabels=saat_araliklari,
        cellLoc="center",
        loc="center"
    )
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(7)

    for (ri, ci), cell in tbl.get_celld().items():
        cell.set_linewidth(0.4)
        cell.set_edgecolor("#BBBBBB")
        if ri == 0:
            cell.set_facecolor(NAVY)
            cell.set_text_props(color="white", fontweight="bold", fontsize=7)
            cell.set_height(0.55)
        elif ci == -1:
            cell.set_facecolor(NAVY)
            cell.set_text_props(color="white", fontweight="bold", fontsize=7)
            cell.set_height(0.45)
        else:
            cell.set_facecolor("#EFF4FB" if ci % 2 == 0 else "#FFFFFF")
            cell.set_text_props(color=NAVY, fontweight="bold")
            cell.set_height(0.45)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=100, bbox_inches="tight")
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode("utf-8")


def xlsx_olustur(veri: list, tarih: str) -> bytes:
    tarih_fmt = datetime.strptime(tarih, "%Y-%m-%d").strftime("%d.%m.%Y")
    NAVY_HEX    = "201F5A"
    NAVY_GRAPHIC = "#201F5A"

    header_fill = PatternFill("solid", start_color=NAVY_HEX, end_color=NAVY_HEX)
    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    wrap_align  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    center_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"), right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),  bottom=Side(style="thin", color="CCCCCC"),
    )

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PTF Verileri"

    for col, w in zip("ABC", [15, 20, 20]):
        ws.column_dimensions[col].width = w

    ws.row_dimensions[1].height = 55
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 15
    ws.merge_cells("B1:C1")
    
    try:
        from openpyxl.drawing.image import Image as XLImage
        xl_logo = XLImage(LOGO_PATH_JPG)
        xl_logo.width  = 148
        xl_logo.height = 45
        xl_logo.anchor = "A1"
        ws.add_image(xl_logo)
    except Exception as e:
        log.warning(f"Logo yüklenemedi (Excel): {e}")
        ws["A1"] = "ALPİNE ENERJİ"
        ws["A1"].font = Font(name="Arial", size=14, bold=True, color=NAVY_HEX)
        ws["A1"].alignment = center_align

    ws["B1"] = f"EPİAŞ Kesinleşmemiş PTF - {tarih_fmt}"
    ws["B1"].font = Font(name="Arial", size=11, bold=True, color=NAVY_HEX)
    ws["B1"].alignment = wrap_align

    ws.row_dimensions[4].height = 22
    for ci, h in enumerate(["Tarih", "Saat Aralığı", "PTF (TL/MWh)"], start=1):
        c = ws.cell(row=4, column=ci, value=h)
        c.font = header_font; c.fill = header_fill; c.alignment = center_align; c.border = thin_border

    for i, row in enumerate(veri, start=5):
        fiyat = float(row["fiyat"])
        bg = "EFF4FB" if i % 2 == 0 else "FFFFFF"
        rf = PatternFill("solid", start_color=bg, end_color=bg)
        bf = Font(name="Arial", size=10, bold=True, color="000000")
        nf = Font(name="Arial", size=10, color="000000")
        c = ws.cell(row=i, column=1, value=tarih_fmt);   c.font=nf; c.alignment=center_align; c.fill=rf; c.border=thin_border
        c = ws.cell(row=i, column=2, value=row["saat"]); c.font=bf; c.alignment=center_align; c.fill=rf; c.border=thin_border
        c = ws.cell(row=i, column=3, value=fiyat);       c.font=bf; c.number_format='#,##0.00'; c.alignment=center_align; c.fill=rf; c.border=thin_border

    # Excel içi grafik
    saatler  = [r["saat_no"] + ":00" for r in veri]
    fiyatlar = [float(r["fiyat"]) for r in veri]
    fig2, ax2 = plt.subplots(figsize=(8, 4))
    ax2.bar(range(len(saatler)), fiyatlar, color=NAVY_GRAPHIC, width=0.7)
    ax2.set_xticks(range(len(saatler)))
    ax2.set_xticklabels(saatler, rotation=45, ha="right", fontsize=8)
    ax2.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: "{:,.0f}".format(x)))
    ax2.grid(axis="y", linestyle="--", alpha=0.3)
    plt.title(f"EPİAŞ Kesinleşmemiş PTF - {tarih_fmt}", fontsize=11, color="#222222", fontweight="bold", pad=25)

    try:
        logo_img2 = imread(LOGO_PATH_PNG)
        logo_ax2  = fig2.add_axes([0.78, 0.88, 0.18, 0.12])
        logo_ax2.imshow(logo_img2)
        logo_ax2.axis("off")
    except Exception as e:
        log.warning(f"Logo yüklenemedi (Excel grafik): {e}")
        fig2.text(0.95, 0.92, "ALPİNE", fontsize=11, fontweight="black", color=NAVY_GRAPHIC, ha="right", va="top")
        fig2.text(0.95, 0.86, "ENERJİ", fontsize=11, fontweight="black", color=NAVY_GRAPHIC, ha="right", va="top")

    plt.tight_layout(rect=[0, 0, 0.9, 0.95])
    ibuf = io.BytesIO()
    fig2.savefig(ibuf, format="png", dpi=100, bbox_inches="tight")
    plt.close(fig2)
    ibuf.seek(0)

    from openpyxl.drawing.image import Image as XLImage
    xl_img = XLImage(ibuf)
    xl_img.anchor = "E5"
    ws.add_image(xl_img)

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
                <td style="vertical-align:middle; text-align:right; width:100px;">
                  <img src="{LOGO_MAIL_SRC}"
                       style="width:100px; height:auto; display:block; margin-left:auto;"
                       alt="Alpine Enerji" />
                </td>
              </tr>
            </table>
          </td>
        </tr>

        <tr>
          <td style="padding:25px 30px;">
            <p style="font-size:15px;">Sayın <b>{musteri_ad}</b>,</p>
            <p style="font-size:15px;">
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
            <p style="font-size:12px; color:#666; border-top:1px solid #eee; padding-top:10px; font-weight:bold;">
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
