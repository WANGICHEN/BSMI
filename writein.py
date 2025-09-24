from docx import Document
from glob import glob
import zipfile
import io
import os
import requests
from io import BytesIO 
from docx.shared import RGBColor

# def write_doc(doc, info):

#     for para in doc.paragraphs:
#         full_text = ''.join(run.text for run in para.runs)
#         replaced = False
#         for key, value in info.items():
#             placeholder = f'{{{key}}}'
#             if placeholder in full_text:
#                 full_text = full_text.replace(placeholder, str(value))
#                 replaced = True
#         if replaced:
#             # 清空原有 runs
#             for run in para.runs:
#                 run.text = ''
#             # 只用一個 run 填回去
#             para.runs[0].text = full_text

#     # 處理所有表格
#     for table in doc.tables:
#         for row in table.rows:
#             for cell in row.cells:
#                 for para in cell.paragraphs:
#                     full_text = ''.join(run.text for run in para.runs)
#                     replaced = False
#                     for key, value in info.items():
#                         placeholder = f'{{{key}}}'
#                         if placeholder in full_text:
#                             full_text = full_text.replace(placeholder, str(value))
#                             replaced = True
#                     if replaced:
#                         for run in para.runs:
#                             run.text = ''
#                             run.font.color.rgb = RGBColor(0, 0, 0)
#                         para.runs[0].text = full_text
#     return doc

def iter_paragraphs(doc):
    """遍歷整份文件的所有段落（含表格、巢狀表格）。"""
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p
                # 巢狀表格
                for sub_table in cell.tables:
                    for r in sub_table.rows:
                        for c in r.cells:
                            for p in c.paragraphs:
                                yield p

def write_doc(doc, info, force_black=True):
    """
    在 run 層級做取代，保留原本樣式；若 force_black=True，凡有替換到的 run 一律改成黑色。
    info 例如：{"name": "王小姐", "date": "2025/09/24"}
    """
    mapping = {f"{{{k}}}": str(v) for k, v in info.items()}

    for para in iter_paragraphs(doc):
        for run in para.runs:
            original = run.text
            new_text = original
            hit = False
            for ph, val in mapping.items():
                if ph in new_text:
                    new_text = new_text.replace(ph, val)
                    hit = True
            if hit:
                run.text = new_text
                if force_black:
                    run.font.color.rgb = RGBColor(0, 0, 0)  # 強制黑色
    return doc

def create_zip(file_dict):
    # file_dict: {"檔名.txt": b"檔案內容", ...}
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname, content in file_dict.items():
            zf.writestr(fname, content)
    buffer.seek(0)
    return buffer



def run_BSMI_doc(info):
    print("run BSMI doc")
    files = {}
    fs = [["00_08.docx", "https://z28856673-my.sharepoint.com/:w:/g/personal/itek_project_i-tek_com_tw/EddnQOquWcRFnEg7TWvy2r0BPAxZon_0AgUEMR8wygTOfA?e=llylt7"],
         ["00_99.docx", "https://z28856673-my.sharepoint.com/:w:/g/personal/itek_project_i-tek_com_tw/EZRXXI9yXRhJuZ1o1WV1iOIBAeD36nOTZe5Ojo5hl7hpmw?e=aKc3ZZ"],
         ["02_01.docx", "https://z28856673-my.sharepoint.com/:w:/g/personal/itek_project_i-tek_com_tw/EQk0sg6ngDxHhPE0pO894Q4BZnkPKc0Y1qYehDwvYPQCdQ?e=leuclc"],
         ["07_01.docx", "https://z28856673-my.sharepoint.com/:w:/g/personal/itek_project_i-tek_com_tw/Ea3dOFrVSFlMtLBljbBtW4oBfJO9g7z8xY8pWIOhA5H-gg?e=V0i1zO"],
         ["外箱標示.docx", "https://z28856673-my.sharepoint.com/:w:/g/personal/itek_project_i-tek_com_tw/ESYvqMkqG9NBlqzSQJQHLWgB9NoCWQLJiPWm-lYVU_pEbQ?e=EeMRfu"]]
    
    for f_name, f in fs:
        information = info.copy()
        if f_name in ["00_08.docx", "外箱標示.docx"]:
            information["series"] = ", " + information["series"]
        download_url = f + ("&download=1" if "?" in f else "?download=1")
            
        headers = {"User-Agent": "Mozilla/5.0"}
        r = requests.get(download_url, headers=headers, allow_redirects=True, timeout=30)
        r.raise_for_status()  # 403/404 會在這裡丟錯
        
        doc = write_doc(Document(BytesIO(r.content)), information)
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        files[f_name] = buf.read()

    zip_buffer = create_zip(files)


    return zip_buffer









