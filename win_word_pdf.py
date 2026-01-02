import os
from config import WIN

try:
    import win32com.client  # type: ignore
except Exception:
    win32com = None

def word_export_pdf_batch(docx_paths: list[str], pdf_paths: list[str]) -> None:
    if not WIN or win32com is None:
        raise RuntimeError("Экспорт DOCX→PDF через Word доступен только на Windows (pywin32).")

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        for docx, pdf in zip(docx_paths, pdf_paths):
            doc = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
            try:
                doc.ExportAsFixedFormat(
                    OutputFileName=os.path.abspath(pdf),
                    ExportFormat=17,
                    OpenAfterExport=False,
                    OptimizeFor=0,
                    CreateBookmarks=1,
                )
            finally:
                doc.Close(False)
    finally:
        word.Quit()