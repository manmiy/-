"""
請求内訳明細表 (PDF) → Excel 転記スクリプト
=============================================
Google Cloud の Document AI (OCR) を使って請求書PDFの明細データを抽出し、
Excel (.xlsx) に転記します。

--------------------------------------------------------------------
■ 事前準備
--------------------------------------------------------------------
1. Google Cloud プロジェクトを作成し、「Document AI API」を有効化する
   https://console.cloud.google.com/apis/library/documentai.googleapis.com

2. Document AI コンソールでプロセッサを作成する
   https://console.cloud.google.com/ai/document-ai/processors
   - 種類は「Document OCR」でOK（表形式の認識精度を上げたい場合は
     「Form Parser」でも可。その場合は下の EXTRACTION_MODE を "form" にする）
   - 作成後に表示される「プロセッサID」を控えておく

3. サービスアカウントを作成し、JSON キーを発行する
   - ロールは「Document AI API のユーザー」があればOK
   - 発行したJSONファイルのパスを環境変数に設定する

     Windows(PowerShell):
       $env:GOOGLE_APPLICATION_CREDENTIALS="C:\\path\\to\\key.json"
     Mac/Linux:
       export GOOGLE_APPLICATION_CREDENTIALS="/path/to/key.json"

4. 必要なライブラリをインストール
     pip install google-cloud-documentai openpyxl

--------------------------------------------------------------------
■ 使い方
--------------------------------------------------------------------
    python invoice_to_excel.py \
        --pdf "4月分.pdf" \
        --project-id "your-gcp-project-id" \
        --location "us" \
        --processor-id "xxxxxxxxxxxxxxxx" \
        --output "4月分_請求内訳.xlsx"

--------------------------------------------------------------------
■ 抽出している列
--------------------------------------------------------------------
発行日 / 伝票No / メーカ / 品番 / 品名 / 数量 / 単位 / 単価 / 金額 / 注文No / 備考
（伝票合計・件名合計の行も別行として転記します）

※ 今回いただいたPDFのレイアウト（1行=発行日 伝票No メーカ 品番 品名 数量 単位 単価 金額 …）
   に合わせて正規表現を組んでいます。実際にOCRにかけた際、スキャンの状態や
   フォントによって空白の入り方が変わることがあるので、抽出結果を1度確認し、
   必要であれば ITEM_PATTERN 等の正規表現を微調整してください。
"""

import argparse
import re
import sys
from pathlib import Path

from google.cloud import documentai_v1 as documentai
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# =====================================================================
# 1. Document AI でPDFをOCR処理
# =====================================================================


def ocr_pdf_with_documentai(pdf_path: str, project_id: str, location: str, processor_id: str):
    """PDFをDocument AIに送信し、認識結果(Documentオブジェクト)を返す"""
    client_options = {"api_endpoint": f"{location}-documentai.googleapis.com"}
    client = documentai.DocumentProcessorServiceClient(client_options=client_options)

    processor_name = client.processor_path(project_id, location, processor_id)

    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()

    raw_document = documentai.RawDocument(content=pdf_bytes, mime_type="application/pdf")
    request = documentai.ProcessRequest(name=processor_name, raw_document=raw_document)

    print("  -> Document AI にリクエスト送信中...")
    result = client.process_document(request=request)
    return result.document


# =====================================================================
# 2. OCRテキストから明細行を正規表現でパース
# =====================================================================

# 1件目の明細行: "04 08 065755 ＪＳＰ MKS30 ミラフォーム３０ｍｍ３種３Ｘ６ 5 枚 1260 6300"
ITEM_PATTERN = re.compile(
    r"^(?P<month>\d{2})\s+(?P<day>\d{2})\s+(?P<denpyo>\d{6})\s+"
    r"(?P<maker>\S+)\s+(?P<hinban>\S+)\s+(?P<hinmei>.+?)\s+"
    r"(?P<qty>-?\d+)\s+(?P<unit>\S+)\s+(?P<price>-?\d+)\s+(?P<amount>-?\d+)"
    r"(?:\s+\((?P<order_no>\d+)\))?"
    r"(?:\s+(?P<note>.+))?$"
)

# 同じ伝票の2行目以降 (発行日・伝票Noが空欄で始まる行)
ITEM_PATTERN_CONT = re.compile(
    r"^(?P<maker>\S+)\s+(?P<hinban>\S+)\s+(?P<hinmei>.+?)\s+"
    r"(?P<qty>-?\d+)\s+(?P<unit>\S+)\s+(?P<price>-?\d+)\s+(?P<amount>-?\d+)"
    r"(?:\s+\((?P<order_no>\d+)\))?"
    r"(?:\s+(?P<note>.+))?$"
)

DENPYO_TOTAL_PATTERN = re.compile(r"^伝票合計\s+(?P<amount>-?\d+)\s*(?:\((?P<order_no>\d+)\))?\s*(?P<note>.+)?$")
KENMEI_TOTAL_PATTERN = re.compile(r"^件名合計\s+(?P<amount>-?\d+)\s*(?:\((?P<order_no>\d+)\))?\s*(?P<note>.+)?$")
TOKUISAKI_TOTAL_PATTERN = re.compile(r"^得意先合計\s+(?P<amount>-?\d+)\s*(?P<note>.+)?$")


def parse_lines(ocr_text: str) -> list[dict]:
    """OCRで得たテキスト(改行区切り)から明細行のリストを作る"""
    rows: list[dict] = []
    current_denpyo = ""
    current_date = ""

    for raw_line in ocr_text.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        # --- 明細行(発行日つき) -----------------------------------
        m = ITEM_PATTERN.match(line)
        if m:
            d = m.groupdict()
            current_date = f"{d['month']}/{d['day']}"
            current_denpyo = d["denpyo"]
            rows.append(_item_row(current_date, current_denpyo, d))
            continue

        # --- 伝票合計行 --------------------------------------------
        m = DENPYO_TOTAL_PATTERN.match(line)
        if m:
            d = m.groupdict()
            rows.append(_total_row(current_date, current_denpyo, "伝票合計", d))
            continue

        # --- 件名合計行 --------------------------------------------
        m = KENMEI_TOTAL_PATTERN.match(line)
        if m:
            d = m.groupdict()
            rows.append(_total_row("", "", "件名合計", d))
            continue

        # --- 得意先合計行 ------------------------------------------
        m = TOKUISAKI_TOTAL_PATTERN.match(line)
        if m:
            d = m.groupdict()
            rows.append(_total_row("", "", "得意先合計", d))
            continue

        # --- 明細行(発行日なし、同伝票の2行目以降) -------------------
        m = ITEM_PATTERN_CONT.match(line)
        if m:
            d = m.groupdict()
            rows.append(_item_row(current_date, current_denpyo, d))
            continue

        # どのパターンにも一致しない行は無視(見出し・空白行など)

    return rows


def _item_row(date: str, denpyo: str, d: dict) -> dict:
    return {
        "発行日": date,
        "伝票No": denpyo,
        "メーカ": d["maker"],
        "品番": d["hinban"],
        "品名": d["hinmei"],
        "数量": int(d["qty"]),
        "単位": d["unit"],
        "単価": int(d["price"]),
        "金額": int(d["amount"]),
        "注文No": d.get("order_no") or "",
        "備考": d.get("note") or "",
        "行種別": "明細",
    }


def _total_row(date: str, denpyo: str, label: str, d: dict) -> dict:
    return {
        "発行日": date,
        "伝票No": denpyo,
        "メーカ": "",
        "品番": "",
        "品名": label,
        "数量": "",
        "単位": "",
        "単価": "",
        "金額": int(d["amount"]),
        "注文No": d.get("order_no") or "",
        "備考": d.get("note") or "",
        "行種別": label,
    }


# =====================================================================
# 3. Excel へ書き出し
# =====================================================================

HEADERS = ["発行日", "伝票No", "メーカ", "品番", "品名", "数量", "単位", "単価", "金額", "注文No", "備考"]
COL_WIDTHS = {"発行日": 8, "伝票No": 10, "メーカ": 16, "品番": 14, "品名": 34,
              "数量": 8, "単位": 6, "単価": 10, "金額": 12, "注文No": 10, "備考": 22}

HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
TOTAL_FILL = PatternFill("solid", fgColor="FCE4D6")
THIN = Side(style="thin", color="AAAAAA")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def write_excel(rows: list[dict], output_path: str, sheet_title: str = "4月分請求内訳") -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title[:31]  # シート名は31文字まで

    # --- ヘッダー行 ---
    for col_idx, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = Font(name="Arial", bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
        cell.border = BORDER

    amount_col = HEADERS.index("金額") + 1
    first_data_row = 2
    detail_rows_for_sum = []

    row_idx = first_data_row
    for row in rows:
        for col_idx, header in enumerate(HEADERS, start=1):
            value = row.get(header, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = Font(name="Arial")
            cell.border = BORDER
            if header in ("数量", "単価", "金額") and value != "":
                cell.number_format = "#,##0"
            if row.get("行種別") != "明細":
                cell.fill = TOTAL_FILL
                cell.font = Font(name="Arial", bold=True)
        if row.get("行種別") == "明細":
            detail_rows_for_sum.append(row_idx)
        row_idx += 1

    # --- 総合計行(明細行のみをSUM式で合計。ハードコードしない) ---
    total_row_idx = row_idx + 1
    ws.cell(row=total_row_idx, column=HEADERS.index("品名") + 1, value="総合計(明細のみ)").font = Font(
        name="Arial", bold=True
    )
    col_letter = get_column_letter(amount_col)
    if detail_rows_for_sum:
        refs = ",".join(f"{col_letter}{r}" for r in detail_rows_for_sum)
        formula = f"=SUM({refs})"
    else:
        formula = 0
    total_cell = ws.cell(row=total_row_idx, column=amount_col, value=formula)
    total_cell.number_format = "#,##0"
    total_cell.font = Font(name="Arial", bold=True)
    total_cell.fill = TOTAL_FILL

    # --- 列幅 ---
    for col_idx, header in enumerate(HEADERS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = COL_WIDTHS.get(header, 12)

    ws.freeze_panes = "A2"
    wb.save(output_path)
    print(f"  -> Excelファイルを保存しました: {output_path}")


# =====================================================================
# 4. メイン処理
# =====================================================================


def main() -> None:
    parser = argparse.ArgumentParser(
        description="請求内訳明細表PDF → Excel 転記 (Google Cloud Document AI使用)"
    )
    parser.add_argument("--pdf", required=True, help="OCR対象のPDFファイルパス")
    parser.add_argument("--project-id", required=True, help="GCPプロジェクトID")
    parser.add_argument("--location", default="us", help="Document AIプロセッサのロケーション (us / eu 等)")
    parser.add_argument("--processor-id", required=True, help="Document AIプロセッサID")
    parser.add_argument("--output", default="請求内訳.xlsx", help="出力するExcelファイル名")
    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    if not pdf_path.exists():
        sys.exit(f"PDFファイルが見つかりません: {pdf_path}")

    print("[1/3] Document AI でOCR処理中...")
    document = ocr_pdf_with_documentai(str(pdf_path), args.project_id, args.location, args.processor_id)

    print("[2/3] 明細行を抽出中...")
    rows = parse_lines(document.text)
    print(f"  -> {len(rows)} 行を抽出しました")
    if not rows:
        print("  !! 1行も抽出できませんでした。OCR結果のテキストを確認し、")
        print("     正規表現(ITEM_PATTERN等)を実際のレイアウトに合わせて調整してください。")
        print("     デバッグ用にOCR全文を確認したい場合は、document.text を print してください。")

    print("[3/3] Excelに書き出し中...")
    write_excel(rows, args.output)


if __name__ == "__main__":
    main()
