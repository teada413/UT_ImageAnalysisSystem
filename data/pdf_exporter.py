"""Excel→PDF変換（COM経由）"""

import os
import time
import win32com.client


def excel_to_pdf(xlsx_path, pdf_path):
    """ExcelファイルをPDFに変換（全シート出力）。

    Excel COMオートメーションを使用。Excelがインストールされている必要がある。

    Args:
        xlsx_path: 入力Excelファイルのパス
        pdf_path: 出力PDFファイルのパス
    """
    xlsx_abs = os.path.abspath(xlsx_path)
    pdf_abs = os.path.abspath(pdf_path)

    # 出力先ディレクトリが存在するか確認
    pdf_dir = os.path.dirname(pdf_abs)
    if pdf_dir and not os.path.exists(pdf_dir):
        os.makedirs(pdf_dir)

    # 既存PDFを事前削除（COMの上書き失敗を回避）
    if os.path.exists(pdf_abs):
        try:
            os.remove(pdf_abs)
        except OSError:
            pass

    excel = None
    wb = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(xlsx_abs, ReadOnly=True)

        # 全シートを選択
        wb.Sheets.Select()

        # PDF出力
        wb.ActiveSheet.ExportAsFixedFormat(
            Type=0,  # xlTypePDF
            Filename=pdf_abs,
            Quality=0,  # xlQualityStandard
            IncludeDocProperties=False,
            IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )
    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass
        # COMオブジェクトの解放を待つ
        time.sleep(0.3)
