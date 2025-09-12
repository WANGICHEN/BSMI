from docx import Document
from glob import glob
import zipfile
import io
import os
from docx.shared import RGBColor

def write_doc(word_path, info):

    doc = Document(word_path)
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



def run_BSMI_doc(word_path, info):
    print("run BSMI doc")
    files = {}
    fs = glob(word_path + '*.docx')
    for f in fs:
        doc = write_doc(f, info)
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        files[os.path.basename(f)] = buf.read()

    zip_buffer = create_zip(files)

    return zip_buffer