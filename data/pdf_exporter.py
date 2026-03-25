"""Excel→PDF変換（COM経由）"""

import os
import win32com.client


def excel_to_pdf(xlsx_path, pdf_path):
    """ExcelファイルをPDFに変換（全シート出力）。

    Excel COMオートメーションを使用。Excelがインストールされている必要がある。

    Args:
        xlsx_path: 入力Excelファイルのパス（絶対パス推奨）
        pdf_path: 出力PDFファイルのパス（絶対パス推奨）
    """
    xlsx_abs = os.path.abspath(xlsx_path)
    pdf_abs = os.path.abspath(pdf_path)

    excel = None
    wb = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(xlsx_abs)

        # 全シートを選択してPDF出力
        # xlTypePDF = 0
        wb.ExportAsFixedFormat(
            Type=0,  # xlTypePDF
            Filename=pdf_abs,
            Quality=0,  # xlQualityStandard
            IncludeDocProperties=True,
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
