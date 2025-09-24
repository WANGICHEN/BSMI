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

from docx.shared import RGBColor

BLACK = RGBColor(0, 0, 0)

def _all_paras(doc):
    for p in doc.paragraphs: yield p
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs: yield p
                for tt in c.tables:
                    for rr in tt.rows:
                        for cc in rr.cells:
                            for p in cc.paragraphs: yield p

def _copy_style(dst, src):
    if src._element.rPr is not None:
        dst._element.rPr = src._element.rPr

def write_doc(doc, mapping, force_black=True):
    # mapping 例： {"{name}": "王小姐", "{date}": "2025/09/24"}
    items = sorted(mapping.items(), key=lambda kv: len(kv[0]), reverse=True)
    for para in _all_paras(doc):
        runs = list(para.runs)
        if not runs: continue
        spans, s = [], 0
        full = "".join(r.text for r in runs)
        if not full: continue
        for r in runs:
            e = s + len(r.text); spans.append((s, e, r)); s = e

        # 找所有匹配（避免重疊）
        occ = [False]*len(full); hits = []
        for ph, val in items:
            start = 0; L = len(ph)
            while True:
                i = full.find(ph, start)
                if i < 0: break
                j = i + L
                if all(not occ[k] for k in range(i, j)):
                    hits.append((i, j, str(val)))
                    for k in range(i, j): occ[k] = True
                start = j
        if not hits: continue
        hits.sort()

        # 清空並重建
        for r in para.runs: r.text = ""
        while para.runs: para._element.remove(para.runs[0]._element)

        cur = 0; h = 0
        while cur < len(full):
            if h < len(hits) and cur == hits[h][0]:
                a,b,val = hits[h]
                # 取樣式來源：placeholder 起點所屬 run
                base = next(r for (s,e,r) in spans if s < a < e or s == a)
                nr = para.add_run(val); _copy_style(nr, base)
                if force_black: nr.font.color.rgb = BLACK
                cur = b; h += 1
            else:
                nxt = hits[h][0] if h < len(hits) else len(full)
                for s,e,r0 in spans:
                    if e <= cur or s >= nxt: continue
                    piece = full[max(s,cur):min(e,nxt)]
                    if piece:
                        nr = para.add_run(piece); _copy_style(nr, r0)
                cur = nxt
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










