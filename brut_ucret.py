"""MUHSGK PDF'lerindeki "Çalışmaya İlişkin Brüt Ücret Tutarı Toplamı" değerlerini Excel'e döker.

Kullanım:
    python brut_ucret.py                 # bu klasördeki tüm PDF'ler
    python brut_ucret.py dosya.pdf ...   # belirli PDF'ler
Çıktı: brut_ucret.xlsx
"""
import re
import sys
from pathlib import Path

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

TC_RE = re.compile(r"^\d{11}$")
DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
AMOUNT_RE = re.compile(r"^\d{1,3}(\.\d{3})*,\d{2}$")
ROW_TOL = 22 # isim satırı TC satırının bu kadar pt üstünde/altında olabilir


def to_float(s):
    return float(s.replace(".", "").replace(",", "."))


def lines(words):
    """Kelimeleri aynı satırdakiler birlikte olacak şekilde gruplar."""
    rows = []
    for w in sorted(words, key=lambda w: (round(w["top"]), w["x0"])):
        if rows and abs(rows[-1][0]["top"] - w["top"]) < 3:
            rows[-1].append(w)
        else:
            rows.append([w])
    return [sorted(r, key=lambda w: w["x0"]) for r in rows]


def parse_pdf(path):
    records, toplam = [], None
    with pdfplumber.open(path) as pdf:
        for page_no, page in enumerate(pdf.pages, 1):
            text = page.extract_text() or ""
            if "Terkin Edilecek Tutara" not in text:
                continue
            words = page.extract_words()
            rows = lines(words)
            for i, row in enumerate(rows):
                texts = [w["text"] for w in row]
                tc_idx = next((k for k, t in enumerate(texts) if TC_RE.match(t)), None)
                if tc_idx is None or tc_idx + 3 >= len(texts):
                    continue
                date, brut, matrah = texts[tc_idx + 1 : tc_idx + 4]
                if not (DATE_RE.match(date) and AMOUNT_RE.match(brut) and AMOUNT_RE.match(matrah)):
                    continue
                tc = row[tc_idx]
                # İsim sütunu: "Çalışanın Adı Soyadı" başlığıyla TC sütunu arasında kalan kelimeler
                name_words = [
                    w for w in words
                    if 177 <= w["x0"] < tc["x0"] - 5 and abs(w["top"] - tc["top"]) <= ROW_TOL
                ]
                name = " ".join(w["text"] for w in sorted(name_words, key=lambda w: (round(w["top"]), w["x0"])))
                # Terkin edilecek tutar: sonraki 6 tutarlık satırın sonuncusu
                terkin = None
                for nxt in rows[i + 1 :]:
                    amts = [w["text"] for w in nxt if AMOUNT_RE.match(w["text"])]
                    if len(amts) == 6:
                        terkin = to_float(amts[-1])
                        break
                records.append({
                    "dosya": path.name, "sayfa": page_no, "ad": name, "tc": tc["text"],
                    "tarih": date, "brut": to_float(brut), "matrah": to_float(matrah), "terkin": terkin,
                })
            m = re.findall(r"^TOPLAM ([\d.]+,\d{2})$", text, re.M)
            if m:
                toplam = to_float(m[-1])  # sayfalardaki TOPLAM kümülatif; son değer genel toplam
    return records, toplam


def write_excel(pdfs, out, log=print):
    """PDF'leri okuyup Excel'i yazar. Tüm PDF'lerin kontrolü tuttuysa True döner."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Brüt Ücretler"
    headers = ["Dosya", "Sayfa", "Ad Soyad", "TC Kimlik No", "İşe Başlama",
               "Çalışmaya İlişkin Brüt Ücret Tutarı Toplamı", "Ücret Matrahı", "Terkin Edilecek Tutar"]
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(bold=True)

    kontrol = wb.create_sheet("Kontrol")
    kontrol.append(["Dosya", "Personel Sayısı", "Brüt Toplam", "Terkin Toplamı (okunan)",
                    "Terkin TOPLAM (PDF)", "Durum"])
    for c in kontrol[1]:
        c.font = Font(bold=True)

    all_ok = True
    row = 1  # son yazılan satır (boş satırlar max_row'u ilerletmediği için elle sayılır)
    for pdf in pdfs:
        records, pdf_toplam = parse_pdf(pdf)
        first = row + 1
        for r in records:
            ws.append([r["dosya"], r["sayfa"], r["ad"], r["tc"], r["tarih"], r["brut"], r["matrah"], r["terkin"]])
        # Her PDF'in kendi toplam satırı (PDF'ler birbirine karışmasın)
        last = row + len(records)
        sums = [f"=SUM({c}{first}:{c}{last})" if records else 0 for c in "FGH"]
        ws.append([f"{pdf.name} TOPLAM", None, None, None, None, *sums])
        for c in ws[ws.max_row]:
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", fgColor="DDEBF7")
        row = last + 2  # toplam satırı + boş satır
        ws.append([])
        brut_sum = round(sum(r["brut"] for r in records), 2)
        terkin_sum = round(sum(r["terkin"] or 0 for r in records), 2)
        ok = pdf_toplam is not None and abs(terkin_sum - pdf_toplam) < 0.01 and all(r["terkin"] is not None for r in records)
        all_ok &= ok
        kontrol.append([pdf.name, len(records), brut_sum, terkin_sum, pdf_toplam,
                        "TAMAM" if ok else "UYUMSUZ - PDF'i elle kontrol edin"])
        kontrol.cell(kontrol.max_row, 6).fill = PatternFill("solid", fgColor="C6EFCE" if ok else "FFC7CE")
        log(f"{pdf.name}: {len(records)} personel, brüt toplam {brut_sum:,.2f} "
              f"({'kontrol TAMAM' if ok else f'UYUMSUZ: okunan terkin {terkin_sum:,.2f} / PDF {pdf_toplam}'})")

    for sheet in (ws, kontrol):
        for col in sheet.columns:
            if col[0].column_letter in "FGH" or sheet is kontrol:
                for c in col[1:]:
                    c.number_format = "#,##0.00"
        for col, width in zip("ABCDEFGH", (45, 7, 28, 14, 12, 22, 16, 16)):
            sheet.column_dimensions[col].width = width
    ws.freeze_panes = "A2"

    wb.save(out)
    log(f"Excel yazıldı: {out}")
    return all_ok


def main():
    here = Path(__file__).parent
    pdfs = [Path(a) for a in sys.argv[1:]] or sorted(here.glob("*.pdf"))
    if not pdfs:
        sys.exit("PDF bulunamadı.")
    if not write_excel(pdfs, here / "brut_ucret.xlsx"):
        sys.exit(1)


if __name__ == "__main__":
    main()
