"""SQLite 数据访问层。

设计说明：
- contacts 表存联系人的主体信息（姓名、分组、公司、备注、是否收藏）；
- contact_methods 表存"一对多"的联系方式，通过 contact_id 外键关联，
  删除联系人时级联删除其所有联系方式（ON DELETE CASCADE）。
这样把"一个联系人多种联系方式"从"表内多列"改为"子表多行"，
既满足需求，也便于后续无限扩展新的联系方式类型。
"""

import os
import sqlite3
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
DB_PATH = os.environ.get("ADDRESS_BOOK_DB") or os.path.join(DATA_DIR, "addressbook.db")

# 允许的联系方式类型（前端下拉框与此保持一致）
METHOD_TYPES = ("phone", "email", "wechat", "qq", "address", "other")
MAX_METHODS_PER_CONTACT = 12

SCHEMA = """
CREATE TABLE IF NOT EXISTS contacts (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    name       TEXT NOT NULL,
    group_name TEXT NOT NULL DEFAULT '未分组',
    company    TEXT NOT NULL DEFAULT '',
    note       TEXT NOT NULL DEFAULT '',
    favorite   INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS contact_methods (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    contact_id INTEGER NOT NULL,
    type       TEXT NOT NULL,
    label      TEXT NOT NULL DEFAULT '',
    value      TEXT NOT NULL,
    sort_order INTEGER NOT NULL DEFAULT 0,
    FOREIGN KEY (contact_id) REFERENCES contacts (id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_methods_contact ON contact_methods (contact_id);
CREATE INDEX IF NOT EXISTS idx_contacts_favorite ON contacts (favorite);
CREATE INDEX IF NOT EXISTS idx_contacts_group ON contacts (group_name);
"""

SEED_CONTACTS = [
    {
        "name": "奶龙", "group_name": "极限编程小队", "company": "青青草原大学 信息工程系",
        "note": "极限编程结对伙伴，负责前端页面。", "favorite": 1,
        "methods": [
            ("phone", "手机", "138 0591 2366"),
            ("email", "邮箱", "nailong@example.com"),
            ("wechat", "微信", "nailong_2024"),
        ],
    },
    {
        "name": "豆包", "group_name": "极限编程小队", "company": "青青草原大学 信息工程系",
        "note": "负责导入导出模块，比较熟 openpyxl。", "favorite": 1,
        "methods": [
            ("phone", "手机", "150 5912 8827"),
            ("qq", "QQ", "842110377"),
            ("address", "宿舍", "羊村宿舍 12 号楼 305"),
        ],
    },
    {
        "name": "熊大", "group_name": "科研小组", "company": "熊洞实验室",
        "note": "带我做实验的前辈，周会时间周四晚。", "favorite": 1,
        "methods": [
            ("phone", "手机", "186 5077 1042"),
            ("email", "学校邮箱", "xiongda@example.com"),
            ("wechat", "微信", "xiongda_lab"),
        ],
    },
    {
        "name": "光头强", "group_name": "朋友", "company": "",
        "note": "老同学，现在在狗熊岭工作。", "favorite": 0,
        "methods": [
            ("phone", "手机", "137 9921 6688"),
            ("wechat", "微信", "guangtouqiang_0912"),
        ],
    },
    {
        "name": "喜羊羊", "group_name": "班级", "company": "青青草原大学 2023 级信息 2 班",
        "note": "班长，请假和材料都找他。", "favorite": 0,
        "methods": [
            ("phone", "手机", "133 0501 7749"),
            ("qq", "QQ", "1203988445"),
            ("email", "邮箱", "xiyangyang@example.com"),
        ],
    },
]


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def get_conn() -> sqlite3.Connection:
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    # timeout：gunicorn 多 worker 同时写库时，等待锁而不是立刻抛 database is locked
    conn = sqlite3.connect(DB_PATH, timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    # WAL 让读操作不阻塞写操作，云服务器上多 worker 场景更稳
    conn.execute("PRAGMA journal_mode = WAL")
    return conn


def init_db(seed: bool | None = None) -> None:
    """建表；若表为空且允许种子数据，则写入演示数据。"""
    with get_conn() as conn:
        conn.executescript(SCHEMA)
        total = conn.execute("SELECT COUNT(*) FROM contacts").fetchone()[0]

    if seed is None:
        seed = os.environ.get("ADDRESS_BOOK_SEED", "1") != "0"
    if seed and total == 0:
        for item in SEED_CONTACTS:
            payload = dict(item)
            payload["group"] = item["group_name"]
            payload["methods"] = [
                {"type": t, "label": l, "value": v} for (t, l, v) in item["methods"]
            ]
            create_contact(payload)


def _row_to_contact(row: sqlite3.Row, methods: list[dict]) -> dict:
    return {
        "id": row["id"],
        "name": row["name"],
        "group": row["group_name"],
        "company": row["company"],
        "note": row["note"],
        "favorite": bool(row["favorite"]),
        "createdAt": row["created_at"],
        "updatedAt": row["updated_at"],
        "methods": methods,
    }


def _load_methods(conn: sqlite3.Connection, contact_ids: list[int]) -> dict[int, list[dict]]:
    if not contact_ids:
        return {}
    placeholders = ",".join("?" * len(contact_ids))
    rows = conn.execute(
        f"""SELECT * FROM contact_methods
            WHERE contact_id IN ({placeholders})
            ORDER BY contact_id, sort_order, id""",
        contact_ids,
    ).fetchall()
    grouped: dict[int, list[dict]] = {}
    for r in rows:
        grouped.setdefault(r["contact_id"], []).append(
            {"id": r["id"], "type": r["type"], "label": r["label"], "value": r["value"]}
        )
    return grouped


def list_contacts(keyword: str = "", favorite_only: bool = False, group: str = "") -> list[dict]:
    sql = "SELECT DISTINCT c.* FROM contacts c LEFT JOIN contact_methods m ON m.contact_id = c.id WHERE 1=1"
    params: list = []
    if favorite_only:
        sql += " AND c.favorite = 1"
    if group:
        sql += " AND c.group_name = ?"
        params.append(group)
    if keyword:
        like = f"%{keyword}%"
        sql += """ AND (
            c.name LIKE ? OR c.group_name LIKE ? OR c.company LIKE ?
            OR c.note LIKE ? OR m.value LIKE ? OR m.label LIKE ?
        )"""
        params.extend([like] * 6)
    sql += " ORDER BY c.favorite DESC, c.updated_at DESC, c.id DESC"

    with get_conn() as conn:
        rows = conn.execute(sql, params).fetchall()
        methods = _load_methods(conn, [r["id"] for r in rows])
    return [_row_to_contact(r, methods.get(r["id"], [])) for r in rows]


def get_contact(contact_id: int) -> dict | None:
    with get_conn() as conn:
        row = conn.execute("SELECT * FROM contacts WHERE id = ?", (contact_id,)).fetchone()
        if row is None:
            return None
        methods = _load_methods(conn, [contact_id])
    return _row_to_contact(row, methods.get(contact_id, []))


def _clean_methods(methods: list[dict]) -> list[tuple[str, str, str]]:
    """校验并清洗联系方式：去空、去重、限制数量。"""
    result: list[tuple[str, str, str]] = []
    seen: set[tuple[str, str]] = set()
    for m in methods or []:
        value = str(m.get("value", "")).strip()
        if not value:
            continue
        mtype = str(m.get("type", "other")).strip() or "other"
        if mtype not in METHOD_TYPES:
            mtype = "other"
        label = str(m.get("label", "")).strip()
        key = (mtype, value)
        if key in seen:
            continue
        seen.add(key)
        result.append((mtype, label, value))
        if len(result) >= MAX_METHODS_PER_CONTACT:
            break
    return result


def create_contact(payload: dict) -> dict:
    name = str(payload.get("name", "")).strip() or "未命名联系人"
    group_name = str(payload.get("group", "")).strip() or "未分组"
    company = str(payload.get("company", "")).strip()
    note = str(payload.get("note", "")).strip()
    favorite = 1 if payload.get("favorite") else 0
    methods = _clean_methods(payload.get("methods", []))
    ts = now_str()

    with get_conn() as conn:
        cur = conn.execute(
            """INSERT INTO contacts (name, group_name, company, note, favorite, created_at, updated_at)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (name, group_name, company, note, favorite, ts, ts),
        )
        contact_id = cur.lastrowid
        conn.executemany(
            "INSERT INTO contact_methods (contact_id, type, label, value, sort_order) VALUES (?, ?, ?, ?, ?)",
            [(contact_id, t, l, v, i) for i, (t, l, v) in enumerate(methods)],
        )
    return get_contact(contact_id)


def update_contact(contact_id: int, payload: dict) -> dict | None:
    if get_contact(contact_id) is None:
        return None
    name = str(payload.get("name", "")).strip() or "未命名联系人"
    group_name = str(payload.get("group", "")).strip() or "未分组"
    company = str(payload.get("company", "")).strip()
    note = str(payload.get("note", "")).strip()
    favorite = 1 if payload.get("favorite") else 0
    methods = _clean_methods(payload.get("methods", []))

    with get_conn() as conn:
        conn.execute(
            """UPDATE contacts SET name = ?, group_name = ?, company = ?, note = ?,
               favorite = ?, updated_at = ? WHERE id = ?""",
            (name, group_name, company, note, favorite, now_str(), contact_id),
        )
        conn.execute("DELETE FROM contact_methods WHERE contact_id = ?", (contact_id,))
        conn.executemany(
            "INSERT INTO contact_methods (contact_id, type, label, value, sort_order) VALUES (?, ?, ?, ?, ?)",
            [(contact_id, t, l, v, i) for i, (t, l, v) in enumerate(methods)],
        )
    return get_contact(contact_id)


def set_favorite(contact_id: int, favorite: bool) -> dict | None:
    if get_contact(contact_id) is None:
        return None
    with get_conn() as conn:
        conn.execute(
            "UPDATE contacts SET favorite = ?, updated_at = ? WHERE id = ?",
            (1 if favorite else 0, now_str(), contact_id),
        )
    return get_contact(contact_id)


def delete_contact(contact_id: int) -> bool:
    with get_conn() as conn:
        cur = conn.execute("DELETE FROM contacts WHERE id = ?", (contact_id,))
        return cur.rowcount > 0


def list_groups() -> list[dict]:
    with get_conn() as conn:
        rows = conn.execute(
            """SELECT group_name AS name, COUNT(*) AS total
               FROM contacts GROUP BY group_name ORDER BY total DESC, group_name"""
        ).fetchall()
    return [{"name": r["name"], "total": r["total"]} for r in rows]


def stats() -> dict:
    with get_conn() as conn:
        total = conn.execute("SELECT COUNT(*) FROM contacts").fetchone()[0]
        fav = conn.execute("SELECT COUNT(*) FROM contacts WHERE favorite = 1").fetchone()[0]
        methods = conn.execute("SELECT COUNT(*) FROM contact_methods").fetchone()[0]
        groups = conn.execute("SELECT COUNT(DISTINCT group_name) FROM contacts").fetchone()[0]
    return {"total": total, "favorite": fav, "methods": methods, "groups": groups}


def find_by_name(name: str) -> list[dict]:
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM contacts WHERE name = ? ORDER BY id", (name.strip(),)
        ).fetchall()
        methods = _load_methods(conn, [r["id"] for r in rows])
    return [_row_to_contact(r, methods.get(r["id"], [])) for r in rows]


def add_methods(contact_id: int, methods: list[dict]) -> dict | None:
    """在已有联系人上追加联系方式（导入合并时使用）。"""
    current = get_contact(contact_id)
    if current is None:
        return None
    merged = list(current["methods"]) + list(methods)
    return update_contact(
        contact_id,
        {
            "name": current["name"],
            "group": current["group"],
            "company": current["company"],
            "note": current["note"],
            "favorite": current["favorite"],
            "methods": merged,
        },
    )
