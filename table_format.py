from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def neutralize_table_style(table):
    # 方式A：直接取消樣式（通常最有效）
    table.style = None

def remove_tc_borders_from_all_cells(table):
    for row in table.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            # 找到並移除 w:tcBorders
            tcBorders = tcPr.find(qn("w:tcBorders"))
            if tcBorders is not None:
                tcPr.remove(tcBorders)

def remove_table_insideV_only(table):
    tbl = table._tbl
    tblPr = tbl.tblPr



def set_table_borders_only(
    table,
    color="000000",
    outside_size=8,
    inside_size=8,
    border_type="single",
    inside=True,
    outside=True
):
    """
    只修改 table 的框線
    size 單位：1/8 pt（8 = 1pt）
    """
    remove_tc_borders_from_all_cells(table)
    neutralize_table_style(table)
    tbl = table._tbl
    tblPr = tbl.tblPr


    # 移除舊的 tblBorders（避免重疊）
    for child in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(child)

    tblBorders = OxmlElement("w:tblBorders")

    def _border(tag, size):
        el = OxmlElement(tag)
        el.set(qn("w:val"), border_type)
        el.set(qn("w:sz"), str(size))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        return el

    if outside:
        print("outside")
        for t in ("top", "left", "bottom", "right"):
            tblBorders.append(_border(f"w:{t}", outside_size))

    if inside:
        tblBorders.append(_border("w:insideH", inside_size))
        tblBorders.append(_border("w:insideV", inside_size))

    tblPr.append(tblBorders)

def _get_or_add(parent, tag):
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        parent.append(el)
    return el

def _set_border_none(el):
    el.set(qn("w:val"), "none")
    el.set(qn("w:sz"), "0")
    el.set(qn("w:space"), "0")
    el.set(qn("w:color"), "auto")

def disable_table_insideV_only(table):
    """只關 table 層級 insideV，不碰外框。"""
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblBorders = tblPr.find(qn("w:tblBorders"))
    if tblBorders is None:
        tblBorders = OxmlElement("w:tblBorders")
        tblPr.append(tblBorders)

    insideV = tblBorders.find(qn("w:insideV"))
    if insideV is None:
        insideV = OxmlElement("w:insideV")
        tblBorders.append(insideV)

    insideV.set(qn("w:val"), "none")
    insideV.set(qn("w:sz"), "0")
    insideV.set(qn("w:space"), "0")
    insideV.set(qn("w:color"), "auto")

def remove_cell_internal_vertical_only_preserve_outer(table):
    """
    只在 cell 層級關掉「內部」垂直邊線：
    - start_col > 0 才把 left 設 none（因為不是最左外框）
    - end_col < last_col 才把 right 設 none（因為不是最右外框）
    外框邊界的 left/right 完全不動，因此外框樣式 100% 保留。
    """
    tbl = table._tbl
    tblGrid = tbl.find(qn("w:tblGrid"))
    if tblGrid is None:
        # 無法精準算欄位時：為了保證不動外框，只做 table insideV 關閉即可
        return

    n_cols = len(tblGrid.findall(qn("w:gridCol")))
    if n_cols <= 1:
        return
    last_col = n_cols - 1

    for tr in tbl.findall(qn("w:tr")):
        col_cursor = 0
        for tc in tr.findall(qn("w:tc")):
            tcPr = tc.find(qn("w:tcPr"))
            if tcPr is None:
                tcPr = OxmlElement("w:tcPr")
                tc.insert(0, tcPr)

            gridSpan = tcPr.find(qn("w:gridSpan"))
            span = int(gridSpan.get(qn("w:val"))) if (gridSpan is not None and gridSpan.get(qn("w:val"))) else 1

            start_col = col_cursor
            end_col = col_cursor + span - 1

            # 只處理「內部」邊界，外框邊界不碰
            if start_col > 0 or end_col < last_col:
                tcBorders = tcPr.find(qn("w:tcBorders"))
                if tcBorders is None:
                    tcBorders = OxmlElement("w:tcBorders")
                    tcPr.append(tcBorders)

                if start_col > 0:
                    left = tcBorders.find(qn("w:left"))
                    if left is None:
                        left = OxmlElement("w:left")
                        tcBorders.append(left)
                    _set_border_none(left)

                if end_col < last_col:
                    right = tcBorders.find(qn("w:right"))
                    if right is None:
                        right = OxmlElement("w:right")
                        tcBorders.append(right)
                    _set_border_none(right)

            col_cursor += span

def set_format(file_name, doc):
    if file_name == "07_01.docx":
        set_table_borders_only(doc.tables[0], outside_size=12, inside_size = 8)
    elif file_name == "02_01.docx":
        print(file_name)
        for table in doc.tables:
            # remove_inside_vertical_by_cell(table)
            remove_cell_internal_vertical_only_preserve_outer(table)
            # remove_inside_vertical(table)
    return doc

