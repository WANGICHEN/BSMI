from docx import Document
from glob import glob
import zipfile
import io
import os
from copy import deepcopy
import requests
from io import BytesIO 
from docx.shared import RGBColor

# def write_doc(doc, info):

#     for para in doc.paragraphs:
#     #     full_text = ''.join(run.text for run in para.runs)
#     #     replaced = False
#     #     for key, value in info.items():
#     #         placeholder = f'{{{key}}}'
#     #         if placeholder in full_text:
#     #             full_text = full_text.replace(placeholder, str(value))
#         #         replaced = True
#         # if replaced: 
#         #     # 清空原有 runs 我這樣的寫法可以怎麼改 
#         #     for run in para.runs: 
#         #         run.text = '' 
#         #         # 只用一個 run 填回去 
#         #     para.runs[0].text = full_text
#         for run in para.runs:
#             for key, value in info.items():
#                 placeholder = f'{{{key}}}'
#                 if placeholder in run.text:
#                     run.text = run.text.replace(placeholder, str(value))
#                     run.font.color.rgb = RGBColor(0, 0, 0)  # 可強制黑色
    
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

from copy import deepcopy
from docx.shared import RGBColor

BLACK = RGBColor(0, 0, 0)

def _all_paras(doc):
    for p in doc.paragraphs: 
        yield p
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs: 
                    yield p
                for tt in c.tables:            # 巢狀表格
                    for rr in tt.rows:
                        for cc in rr.cells:
                            for p in cc.paragraphs: 
                                yield p

def _copy_rpr(dst_run, src_run):
    """把來源 run 的樣式複製到新 run（不直接賦值，避免 lxml parent 衝突）。"""
    src = src_run._element.rPr
    if src is None: 
        return
    el = dst_run._element
    if el.rPr is not None:
        el.remove(el.rPr)
    el.insert(0, deepcopy(src))

def write_doc(doc, mapping, force_black=True, prefer="first"):
    """
    mapping: 例如 {"{name}": "王小姐", "{date}": "2025/09/24"}
    force_black: True=替換文字設為黑色；False=沿用來源樣式顏色
    prefer: "first"=沿用 placeholder 起點 run 樣式；"last"=沿用終點 run 樣式
    """
    # 先長字串優先，避免互相覆蓋
    items = sorted(mapping.items(), key=lambda kv: len(kv[0]), reverse=True)

    for para in _all_paras(doc):
        runs = list(para.runs)
        if not runs:
            continue

        # 攤平成整段文字與 run 對應區間
        full = "".join(r.text for r in runs)
        if not full:
            continue
        spans, s = [], 0
        for r in runs:
            e = s + len(r.text)
            spans.append((s, e, r))
            s = e

        # 找不重疊的所有匹配
        used = [False] * len(full)
        hits = []  # (start, end, value, base_run)
        for ph, val in items:
            i, L = 0, len(ph)
            while True:
                j = full.find(ph, i)
                if j < 0:
                    break
                k = j + L
                if all(not used[t] for t in range(j, k)):
                    involved = [r for (a, b, r) in spans if not (b <= j or a >= k)]
                    base = involved[0] if prefer == "first" else involved[-1]
                    hits.append((j, k, str(val), base))
                    for t in range(j, k):
                        used[t] = True
                i = k
        if not hits:
            continue
        hits.sort(key=lambda x: x[0])

        # 清空並依序重建（替換片段→新 run；其他片段→按原 run 切片重建）
        for r in para.runs:
            r.text = ""
        while para.runs:
            para._element.remove(para.runs[0]._element)

        cur, h = 0, 0
        while cur < len(full):
            if h < len(hits) and cur == hits[h][0]:
                a, b, val, base = hits[h]
                nr = para.add_run(val)
                _copy_rpr(nr, base)
                if force_black:
                    nr.font.color.rgb = BLACK
                cur = b
                h += 1
            else:
                nxt = hits[h][0] if h < len(hits) else len(full)
                for a, b, r0 in spans:
                    if b <= cur or a >= nxt:
                        continue
                    piece = full[max(a, cur):min(b, nxt)]
                    if piece:
                        nr = para.add_run(piece)
                        _copy_rpr(nr, r0)
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
            information["{series}"] = ", " + information["{series}"]
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




















