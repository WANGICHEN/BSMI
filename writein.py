from docx import Document
from glob import glob
import zipfile
import io
import os
import requests
from docx.shared import RGBColor

def write_doc(doc, info):

    for para in doc.paragraphs:
        full_text = ''.join(run.text for run in para.runs)
        replaced = False
        for key, value in info.items():
            placeholder = f'{{{key}}}'
            if placeholder in full_text:
                full_text = full_text.replace(placeholder, str(value))
                replaced = True
        if replaced:
            # 清空原有 runs
            for run in para.runs:
                run.text = ''
            # 只用一個 run 填回去
            para.runs[0].text = full_text

    # 處理所有表格
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    full_text = ''.join(run.text for run in para.runs)
                    replaced = False
                    for key, value in info.items():
                        placeholder = f'{{{key}}}'
                        if placeholder in full_text:
                            full_text = full_text.replace(placeholder, str(value))
                            replaced = True
                    if replaced:
                        for run in para.runs:
                            run.text = ''
                            run.font.color.rgb = RGBColor(0, 0, 0)
                        para.runs[0].text = full_text
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
    fs = ["https://z28856673-my.sharepoint.com/:w:/g/personal/itek_project_i-tek_com_tw/EddnQOquWcRFnEg7TWvy2r0BPAxZon_0AgUEMR8wygTOfA?e=llylt7",
         "https://z28856673-my.sharepoint.com/:w:/g/personal/itek_project_i-tek_com_tw/EZRXXI9yXRhJuZ1o1WV1iOIBAeD36nOTZe5Ojo5hl7hpmw?e=aKc3ZZ",
         "https://z28856673-my.sharepoint.com/:w:/g/personal/itek_project_i-tek_com_tw/EQk0sg6ngDxHhPE0pO894Q4BZnkPKc0Y1qYehDwvYPQCdQ?e=leuclc",
         "https://z28856673-my.sharepoint.com/:w:/g/personal/itek_project_i-tek_com_tw/Ea3dOFrVSFlMtLBljbBtW4oBfJO9g7z8xY8pWIOhA5H-gg?e=V0i1zO",
         "https://z28856673-my.sharepoint.com/:w:/g/personal/itek_project_i-tek_com_tw/ESYvqMkqG9NBlqzSQJQHLWgB9NoCWQLJiPWm-lYVU_pEbQ?e=EeMRfu"]
    
    for f in fs:
        download_url = f + ("&download=1" if "?" in f else "?download=1")
            
        headers = {"User-Agent": "Mozilla/5.0"}
        r = requests.get(download_url, headers=headers, allow_redirects=True, timeout=30)
        r.raise_for_status()  # 403/404 會在這裡丟錯
        
        doc = write_doc(Document(BytesIO(r.content)), info)
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        files[os.path.basename(f)] = buf.read()

    zip_buffer = create_zip(files)


    return zip_buffer


