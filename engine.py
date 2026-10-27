import os
import html
import uuid
import fitz  # PyMuPDF
import img2pdf
from PIL import Image
from pdf2docx import Converter
import pandas as pd
import win32com.client
import pythoncom

SUPPORTED_IMAGES = {'.jpg', '.jpeg', '.png', '.webp', '.bmp'}
WORD_EXTS = {'.docx', '.doc', '.rtf', '.txt'}
PPT_EXTS = {'.pptx', '.ppt'}
EXCEL_EXTS = {'.xlsx', '.xls'}
TABULAR_EXTS = {'.csv', '.xlsx', '.xls'}
HTML_EXTS = {'.html', '.htm'}
CODE_EXTS = {'.py', '.c', '.cpp', '.h', '.hpp', '.sql', '.js', '.css', '.json', '.md', '.matlab', '.m'}


def convert_files(
    file_paths: list[str],
    target_format: str,
    merge_files: bool = True,
    custom_output: str = None
) -> list[str]:
    target = target_format.lower().strip('.')
    output_files = []

    valid_paths = [os.path.abspath(p) for p in file_paths if os.path.isfile(p)]
    if not valid_paths:
        raise ValueError("No valid files selected")

    # =========================================================================
    # CASE 1: MERGING MULTIPLE FILES (2+ files selected with Merge checked)
    # =========================================================================
    if merge_files and len(valid_paths) > 1:
        first_dir = os.path.dirname(valid_paths[0])
        if custom_output and not os.path.isdir(custom_output):
            out_path = custom_output if custom_output.lower().endswith(f".{target}") else f"{custom_output}.{target}"
        else:
            save_dir = custom_output if (custom_output and os.path.isdir(custom_output)) else first_dir
            out_path = os.path.join(save_dir, f"merged_output.{target}")

        # 1A. Merge ANY combination -> 1 PDF
        if target == 'pdf':
            _merge_any_to_pdf(valid_paths, out_path)
            return [out_path]

        # 1B. Merge ANY combination -> 1 DOCX
        elif target == 'docx':
            temp_master_pdf = os.path.join(first_dir, f"_temp_master_{uuid.uuid4().hex[:6]}.pdf")
            try:
                _merge_any_to_pdf(valid_paths, temp_master_pdf)
                cv = Converter(temp_master_pdf)
                cv.convert(out_path, start=0, end=None)
                cv.close()
            finally:
                if os.path.exists(temp_master_pdf):
                    try:
                        os.remove(temp_master_pdf)
                    except Exception:
                        pass
            return [out_path]

        # 1C. Merge Multiple CSV / Excel files -> 1 Multi-Sheet XLSX
        elif target == 'xlsx' and all(os.path.splitext(p)[1].lower() in TABULAR_EXTS for p in valid_paths):
            with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
                for idx, p in enumerate(valid_paths):
                    ext = os.path.splitext(p)[1].lower()
                    df = pd.read_csv(p) if ext == '.csv' else pd.read_excel(p)
                    sheet_name = f"{idx+1}_{os.path.splitext(os.path.basename(p))[0]}"[:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            return [out_path]

        # 1D. Merge Multiple CSV / Excel files -> 1 Combined CSV
        elif target == 'csv' and all(os.path.splitext(p)[1].lower() in TABULAR_EXTS for p in valid_paths):
            dfs = [
                pd.read_csv(p) if os.path.splitext(p)[1].lower() == '.csv' else pd.read_excel(p)
                for p in valid_paths
            ]
            pd.concat(dfs, ignore_index=True).to_csv(out_path, index=False)
            return [out_path]

    # =========================================================================
    # CASE 2: SINGLE FILE OR UN-MERGED BATCH CONVERSIONS
    # =========================================================================
    for abs_path in valid_paths:
        base, ext = os.path.splitext(abs_path)
        ext = ext.lower()

        if custom_output:
            if len(valid_paths) == 1 and not os.path.isdir(custom_output):
                out_path = custom_output if custom_output.lower().endswith(f".{target}") else f"{custom_output}.{target}"
            else:
                filename = os.path.basename(base)
                out_path = os.path.join(custom_output, f"{filename}_converted.{target}")
        else:
            out_path = f"{base}_converted.{target}"

        # 0. Safe copy if 1 PDF is converted to PDF
        if ext == '.pdf' and target == 'pdf':
            doc = fitz.open(abs_path)
            doc.save(out_path)
            doc.close()
            output_files.append(out_path)

        # 1. Image -> PDF
        elif ext in SUPPORTED_IMAGES and target == 'pdf':
            _images_to_pdf([abs_path], out_path)
            output_files.append(out_path)

        # 2. Word / PowerPoint / Excel / TXT -> PDF
        elif (ext in WORD_EXTS or ext in PPT_EXTS or ext in EXCEL_EXTS) and target == 'pdf':
            _office_to_pdf(abs_path, out_path, ext)
            output_files.append(out_path)

        # 3. HTML (.html, .htm) -> Rendered Webpage PDF
        elif ext in HTML_EXTS and target == 'pdf':
            _html_or_code_to_pdf(abs_path, out_path, is_code=False)
            output_files.append(out_path)

        # 4. Source Code (.py, .c, .cpp, .sql, .js, etc.) -> Formatted Code PDF
        elif ext in CODE_EXTS and target == 'pdf':
            _html_or_code_to_pdf(abs_path, out_path, is_code=True)
            output_files.append(out_path)

        # 5. Excel / PPT / Image / HTML / Code -> DOCX (via temporary PDF bridge)
        elif (ext in EXCEL_EXTS or ext in PPT_EXTS or ext in SUPPORTED_IMAGES or ext in HTML_EXTS or ext in CODE_EXTS) and target == 'docx':
            temp_pdf = f"{base}_temp_{uuid.uuid4().hex[:6]}.pdf"
            try:
                convert_files([abs_path], 'pdf', merge_files=False, custom_output=temp_pdf)
                cv = Converter(temp_pdf)
                cv.convert(out_path, start=0, end=None)
                cv.close()
            finally:
                if os.path.exists(temp_pdf):
                    try:
                        os.remove(temp_pdf)
                    except Exception:
                        pass
            output_files.append(out_path)

        # 6. Image -> Image (PNG, JPG, WEBP, BMP)
        elif ext in SUPPORTED_IMAGES and f".{target}" in SUPPORTED_IMAGES:
            with Image.open(abs_path) as img:
                if target in ['jpg', 'jpeg'] and img.mode in ('RGBA', 'P'):
                    img = img.convert('RGB')
                img.save(out_path)
            output_files.append(out_path)

        # 7. PDF -> DOCX
        elif ext == '.pdf' and target == 'docx':
            cv = Converter(abs_path)
            cv.convert(out_path, start=0, end=None)
            cv.close()
            output_files.append(out_path)

        # 8. PDF -> Images (PNG / JPG)
        elif ext == '.pdf' and target in ['png', 'jpg', 'jpeg']:
            with fitz.open(abs_path) as doc:
                for page_num in range(len(doc)):
                    page = doc.load_page(page_num)
                    pix = page.get_pixmap(dpi=200)
                    img_out = out_path if len(doc) == 1 else f"{os.path.splitext(out_path)[0]}_p{page_num + 1}.{target}"
                    pix.save(img_out)
                    output_files.append(img_out)

        # 9. CSV <-> XLSX / PDF
        elif ext == '.csv' and target == 'xlsx':
            pd.read_csv(abs_path).to_excel(out_path, index=False)
            output_files.append(out_path)
        elif ext in EXCEL_EXTS and target == 'csv':
            pd.read_excel(abs_path).to_csv(out_path, index=False)
            output_files.append(out_path)
        elif ext == '.csv' and target == 'pdf':
            temp_xlsx = f"{base}_temp_{uuid.uuid4().hex[:6]}.xlsx"
            try:
                pd.read_csv(abs_path).to_excel(temp_xlsx, index=False)
                _office_to_pdf(temp_xlsx, out_path, '.xlsx')
            finally:
                if os.path.exists(temp_xlsx):
                    try:
                        os.remove(temp_xlsx)
                    except Exception:
                        pass
            output_files.append(out_path)

        else:
            raise ValueError(f"Unsupported: {ext} -> .{target}")

    return output_files


def _html_or_code_to_pdf(input_path: str, output_pdf: str, is_code: bool = False):
    """Renders HTML files as visual webpages, or wraps code files in clean monospace formatting."""
    with open(input_path, 'r', encoding='utf-8', errors='ignore') as f:
        raw_text = f.read()

    if is_code:
        escaped_code = html.escape(raw_text)
        filename = html.escape(os.path.basename(input_path))
        html_str = f"""
        <html>
        <head>
            <style>
                body {{ font-family: Consolas, 'Courier New', monospace; font-size: 9.5pt; margin: 24px; }}
                h3 {{ font-family: 'Segoe UI', sans-serif; color: #333; border-bottom: 1px solid #ccc; padding-bottom: 4px; }}
                pre {{ background-color: #f6f8fa; padding: 12px; border: 1px solid #d0d7de; white-space: pre-wrap; word-wrap: break-word; }}
            </style>
        </head>
        <body>
            <h3>File: {filename}</h3>
            <pre>{escaped_code}</pre>
        </body>
        </html>
        """
    else:
        html_str = raw_text

    # Open as an HTML document stream, convert stream to PDF bytes, and save
    with fitz.open(stream=html_str.encode('utf-8'), filetype="html") as html_doc:
        pdf_bytes = html_doc.convert_to_pdf()
        with fitz.open("pdf", pdf_bytes) as pdf_doc:
            pdf_doc.save(output_pdf)


def _merge_any_to_pdf(file_paths: list[str], out_path: str):
    """Merges any combination of PDFs, Images, Word, Excel, PPT, HTML, and Code into 1 PDF."""
    if all(os.path.splitext(p)[1].lower() in SUPPORTED_IMAGES for p in file_paths):
        _images_to_pdf(file_paths, out_path)
        return

    merged_doc = fitz.open()
    temp_pdfs = []
    try:
        for p in file_paths:
            abs_p = os.path.abspath(p)
            ext = os.path.splitext(abs_p)[1].lower()

            if ext == '.pdf':
                with fitz.open(abs_p) as src_pdf:
                    merged_doc.insert_pdf(src_pdf)
            else:
                temp_pdf = f"{abs_p}_temp_{uuid.uuid4().hex[:6]}.pdf"
                convert_files([abs_p], 'pdf', merge_files=False, custom_output=temp_pdf)
                temp_pdfs.append(temp_pdf)
                with fitz.open(temp_pdf) as src_pdf:
                    merged_doc.insert_pdf(src_pdf)

        merged_doc.save(out_path)
    finally:
        merged_doc.close()
        for t in temp_pdfs:
            if os.path.exists(t):
                try:
                    os.remove(t)
                except Exception:
                    pass


def _office_to_pdf(input_path: str, output_pdf: str, ext: str):
    pythoncom.CoInitialize()
    app = None
    try:
        if ext in WORD_EXTS:
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            doc = app.Documents.Open(input_path, ReadOnly=True)
            doc.SaveAs(output_pdf, FileFormat=17)
            doc.Close(False)
        elif ext in PPT_EXTS:
            app = win32com.client.DispatchEx("PowerPoint.Application")
            deck = app.Presentations.Open(input_path, ReadOnly=True, Untitled=False, WithWindow=False)
            deck.SaveAs(output_pdf, 32)
            deck.Close()
        elif ext in EXCEL_EXTS:
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            wb = app.Workbooks.Open(input_path, ReadOnly=True)
            for ws in wb.Worksheets:
                ws.PageSetup.Zoom = False
                ws.PageSetup.FitToPagesWide = 1
                ws.PageSetup.FitToPagesTall = False
            wb.ExportAsFixedFormat(0, output_pdf)
            wb.Close(False)
    finally:
        if app:
            app.Quit()
        pythoncom.CoUninitialize()


def _images_to_pdf(img_paths: list[str], output_pdf: str):
    clean_bytes = []
    for p in img_paths:
        with Image.open(p) as img:
            if img.mode in ("RGBA", "P") or p.lower().endswith(('.png', '.webp')):
                img = img.convert("RGB")
                temp = f"{p}_temp_{uuid.uuid4().hex[:6]}.jpg"
                img.save(temp, "JPEG", quality=95)
                with open(temp, "rb") as f:
                    clean_bytes.append(f.read())
                os.remove(temp)
            else:
                with open(p, "rb") as f:
                    clean_bytes.append(f.read())

    with open(output_pdf, "wb") as f:
        f.write(img2pdf.convert(clean_bytes))