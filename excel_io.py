"""Excel 导入 / 导出模块（openpyxl）。

导出格式（一人一行，一条信息一列）：
    姓名 | 分组 | 收藏 | 公司 | 联系方式1类型 | 联系方式1内容 | ... | 联系方式6类型 | 联系方式6内容 | 备注

- 类型列写中文类型名，若该条联系方式带标签（例如"办公室"），写成 `手机/办公室`；
- 导入时同时兼容上面这种"槽位格式"和"手机号 / 邮箱 / 微信 / QQ / 地址 各占一列"的简易格式，
  这样用 Excel 手写一张表也能直接导进来。
"""

import re
from io import BytesIO

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

SLOT_COUNT = 6
IMPORT_MAX_ROWS = 2000

TYPE_LABELS = {
    "phone": "手机",
    "email": "邮箱",
    "wechat": "微信",
    "qq": "QQ",
    "address": "地址",
    "other": "其他",
}
LABEL_TO_TYPE = {
    "手机": "phone", "电话": "phone", "手机号": "phone", "手机号码": "phone",
    "联系电话": "phone", "tel": "phone", "phone": "phone", "mobile": "phone",
    "邮箱": "email", "邮件": "email", "电子邮箱": "email", "email": "email", "mail": "email",
    "微信": "wechat", "微信号": "wechat", "wechat": "wechat", "weixin": "wechat",
    "qq": "qq", "qq号": "qq",
    "地址": "address", "住址": "address", "通讯地址": "address", "address": "address",
    "其他": "other", "其它": "other", "other": "other",
}

BASE_COLUMNS = ["姓名", "分组", "收藏", "公司"]

HEADER_FILL = PatternFill("solid", fgColor="4F46E5")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
THIN = Side(style="thin", color="D9DCE7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _slot_headers() -> list[str]:
    headers: list[str] = []
    for i in range(1, SLOT_COUNT + 1):
        headers.extend([f"联系方式{i}类型", f"联系方式{i}内容"])
    return headers


EXPORT_HEADERS = BASE_COLUMNS + _slot_headers() + ["备注"]
SLOT_PATTERN = re.compile(r"^联系方式(\d+)(类型|内容|标签|值)$")


def _type_cell(method: dict) -> str:
    label = TYPE_LABELS.get(method["type"], "其他")
    extra = (method.get("label") or "").strip()
    return f"{label}/{extra}" if extra else label


def _style_sheet(ws, headers: list[str], widths: list[int]) -> None:
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(ws.max_row, 1)}"
    for col, (header, width) in enumerate(zip(headers, widths), start=1):
        cell = ws.cell(row=1, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[1].height = 24


def export_contacts(contacts: list[dict]) -> BytesIO:
    """把联系人列表导出为 xlsx 字节流。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "通讯录"
    ws.append(EXPORT_HEADERS)

    for c in contacts:
        row = [c["name"], c["group"], "是" if c["favorite"] else "否", c.get("company", "")]
        methods = c.get("methods", [])
        for i in range(SLOT_COUNT):
            if i < len(methods):
                row.extend([_type_cell(methods[i]), methods[i]["value"]])
            else:
                row.extend(["", ""])
        row.append(c.get("note", ""))
        ws.append(row)

    widths = [12, 12, 6, 22] + [11, 22] * SLOT_COUNT + [32]
    _style_sheet(ws, EXPORT_HEADERS, widths)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(EXPORT_HEADERS)):
        for cell in row:
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=False)

    # 第二个 sheet：格式说明，导出的文件自己就能解释自己
    guide = wb.create_sheet("说明")
    guide["A1"] = "通讯录导出说明"
    guide["A1"].font = Font(bold=True, size=14, color="4F46E5")
    lines = [
        "",
        "1. 一个联系人占一行，一条信息占一列，可以直接在 Excel 里编辑后重新导入。",
        f"2. 每个联系人最多支持 {SLOT_COUNT} 条联系方式，类型可选：手机 / 邮箱 / 微信 / QQ / 地址 / 其他。",
        "3. 类型列可以写成 `手机/办公室` 这种形式，斜杠后面是该条联系方式的标签；只写类型也可以。",
        "4. 收藏列填「是」或「否」。",
        "5. 导入时若某条联系方式留空，该行不报错，只跳过这一格。",
        f"6. 单次导入上限 {IMPORT_MAX_ROWS} 行。",
    ]
    for i, text in enumerate(lines, start=2):
        guide.cell(row=i, column=1, value=text)
    guide.column_dimensions["A"].width = 96

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def build_template() -> BytesIO:
    """生成导入模板：表头 + 两行示例数据。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "通讯录"
    ws.append(EXPORT_HEADERS)
    ws.append(["小灰灰", "极限编程小队", "是", "青青草原大学", "手机/常用", "13800001111",
               "邮箱", "xiaohuihui@example.com", "微信", "xiaohuihui_wx", "", "", "", "", "", "", "示例行，可删除"])
    ws.append(["红太狼", "朋友", "否", "", "手机", "13900002222", "QQ", "123456789", "", "",
               "", "", "", "", "", "", "只有两条联系方式"])
    widths = [12, 12, 6, 22] + [11, 22] * SLOT_COUNT + [32]
    _style_sheet(ws, EXPORT_HEADERS, widths)

    ws2 = wb.create_sheet("填写说明")
    notes = [
        ("填写说明", True),
        ("", False),
        ("① 第一行表头请勿修改，从第二行开始填写联系人。", False),
        ("② 姓名必填，其余列可留空。", False),
        ("③ 联系方式按「类型 + 内容」成对填写，最多 6 组，多余的组留空即可。", False),
        ("④ 类型可填：手机 / 邮箱 / 微信 / QQ / 地址 / 其他；也可写成「手机/办公室」，斜杠后为标签。", False),
        ("⑤ 收藏列填「是」或「否」。", False),
        ("⑥ 本模板同样支持「手机号、邮箱、微信、QQ、地址」各占一列的简易表头格式。", False),
    ]
    for i, (text, bold) in enumerate(notes, start=1):
        cell = ws2.cell(row=i, column=1, value=text)
        cell.font = Font(bold=bold, size=14 if bold else 11, color="4F46E5" if bold else "333333")
    ws2.column_dimensions["A"].width = 100

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def _norm(value) -> str:
    """表头归一化：去空白、全角空格，转小写。"""
    if value is None:
        return ""
    return str(value).replace(" ", "").replace("\u3000", "").replace("\t", "").strip().lower()


def _cell_text(value) -> str:
    """把单元格内容转成字符串，避免手机号被 Excel 存成浮点数后变成科学计数法。"""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "是" if value else "否"
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else str(value)
    if isinstance(value, int):
        return str(value)
    return str(value).strip()


def _split_multi(text: str) -> list[str]:
    parts = re.split(r"[\n\r；;、]+", text)
    return [p.strip() for p in parts if p.strip()]


def _parse_favorite(text: str) -> bool:
    return _norm(text) in {"是", "y", "yes", "true", "1", "收藏", "★"}


def _build_column_map(headers: list) -> dict:
    """解析表头，返回 {字段: 列下标} 以及槽位映射。"""
    mapping: dict = {"name": None, "group": None, "favorite": None, "company": None, "note": None}
    slots: dict[int, dict] = {}
    type_columns: list[tuple[int, str]] = []
    label_columns: dict[int, int] = {}

    aliases = {
        "name": {"姓名", "名字", "name", "联系人"},
        "group": {"分组", "分组/标签", "分组标签", "标签", "组别", "group"},
        "favorite": {"收藏", "是否收藏", "星标", "favorite"},
        "company": {"公司", "单位", "公司/单位", "company"},
        "note": {"备注", "remark", "note", "说明"},
    }

    for idx, raw in enumerate(headers):
        key = _norm(raw)
        if not key:
            continue
        matched = False
        for field, names in aliases.items():
            if key in names and mapping[field] is None:
                mapping[field] = idx
                matched = True
                break
        if matched:
            continue

        slot = SLOT_PATTERN.match(key)
        if slot:
            number, part = int(slot.group(1)), slot.group(2)
            slots.setdefault(number, {})[part] = idx
            continue

        if key in LABEL_TO_TYPE:
            type_columns.append((idx, LABEL_TO_TYPE[key]))

    for number, cols in slots.items():
        if "标签" in cols and "内容" in cols:
            label_columns[cols["内容"]] = cols["标签"]

    return {"fields": mapping, "slots": slots, "typeColumns": type_columns, "labelColumns": label_columns}


def _find_header_row(rows: list[list]) -> tuple[int, list] | tuple[None, None]:
    for i, row in enumerate(rows[:6]):
        for cell in row:
            if _norm(cell) in {"姓名", "name", "联系人"}:
                return i, row
    return None, None


def parse_workbook(stream) -> tuple[list[dict], list[dict]]:
    """解析上传的 Excel，返回 (联系人列表, 错误列表)。"""
    wb = load_workbook(filename=stream, read_only=True, data_only=True)
    sheet = None
    for ws in wb.worksheets:
        if ws.title in ("通讯录", "Sheet1", "联系人"):
            sheet = ws
            break
    if sheet is None:
        sheet = wb.worksheets[0]

    raw_rows = [list(r) for r in sheet.iter_rows(values_only=True)]
    if not raw_rows:
        return [], [{"row": 0, "message": "文件是空的，没有读到任何内容"}]

    header_index, headers = _find_header_row(raw_rows)
    if header_index is None:
        return [], [{"row": 1, "message": "没找到表头，请确保表格中包含「姓名」列"}]

    column_map = _build_column_map(headers)
    if column_map["fields"]["name"] is None:
        return [], [{"row": header_index + 1, "message": "表头缺少「姓名」列"}]

    contacts: list[dict] = []
    errors: list[dict] = []
    seen_keys: set[str] = set()
    data_rows = raw_rows[header_index + 1:]

    if len(data_rows) > IMPORT_MAX_ROWS:
        return [], [{"row": 0, "message": f"单次最多导入 {IMPORT_MAX_ROWS} 行，当前 {len(data_rows)} 行"}]

    fields = column_map["fields"]

    def pick(row: list, idx: int | None) -> str:
        if idx is None or idx >= len(row):
            return ""
        return _cell_text(row[idx])

    for offset, row in enumerate(data_rows):
        row_number = header_index + 2 + offset
        if all(_cell_text(v) == "" for v in row):
            continue  # 整行为空，跳过

        name = pick(row, fields["name"])
        if not name:
            errors.append({"row": row_number, "message": "姓名为空，已跳过该行"})
            continue

        methods: list[dict] = []
        seen_methods: set[tuple[str, str]] = set()

        def push(mtype: str, label: str, value: str) -> None:
            value = value.strip()
            if not value:
                return
            key = (mtype, value)
            if key in seen_methods:
                return
            seen_methods.add(key)
            methods.append({"type": mtype, "label": label, "value": value})

        for number in sorted(column_map["slots"]):
            cols = column_map["slots"][number]
            if "内容" not in cols:
                continue
            type_text = pick(row, cols.get("类型"))
            value_text = pick(row, cols["内容"])
            if not value_text:
                continue
            type_name, _, label = type_text.replace("／", "/").partition("/")
            mtype = LABEL_TO_TYPE.get(_norm(type_name), "other")
            if not type_name.strip():
                mtype = "other"
            push(mtype, label.strip(), value_text)

        for idx, mtype in column_map["typeColumns"]:
            values = _split_multi(pick(row, idx))
            for value in values:
                push(mtype, "", value)

        if not methods:
            errors.append({"row": row_number, "message": f"「{name}」没有任何联系方式，已按空联系人导入"})

        key = f"{name}|{methods[0]['value'] if methods else ''}"
        if key in seen_keys:
            errors.append({"row": row_number, "message": f"「{name}」在文件中重复出现，已跳过"})
            continue
        seen_keys.add(key)

        contacts.append(
            {
                "name": name,
                "group": pick(row, fields["group"]) or "未分组",
                "favorite": _parse_favorite(pick(row, fields["favorite"])),
                "company": pick(row, fields["company"]),
                "note": pick(row, fields["note"]),
                "methods": methods,
            }
        )

    return contacts, errors[:20]
