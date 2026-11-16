"""通讯录 Web 应用入口（Flask）。

运行：
    python app.py            # 默认 http://127.0.0.1:5000
    PORT=8000 python app.py  # 指定端口（云服务器部署时使用）
"""

import io
import os
from datetime import datetime

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.exceptions import HTTPException

import db
import excel_io

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # 上传文件上限 10MB
app.json.ensure_ascii = False

db.init_db()  # 首次启动自动建表并写入演示数据

ALLOWED_SUFFIX = (".xlsx", ".xlsm")


@app.errorhandler(HTTPException)
def handle_http_error(err: HTTPException):
    if request.path.startswith("/api/"):
        return jsonify({"ok": False, "message": err.description}), err.code
    return err


@app.route("/")
def index():
    return render_template("index.html")


@app.get("/api/health")
def health():
    return jsonify({"ok": True, "time": db.now_str()})


@app.get("/api/contacts")
def api_list_contacts():
    keyword = (request.args.get("keyword") or "").strip()
    favorite_only = request.args.get("favorite") in ("1", "true", "yes")
    group = (request.args.get("group") or "").strip()
    contacts = db.list_contacts(keyword=keyword, favorite_only=favorite_only, group=group)
    return jsonify({"ok": True, "data": contacts, "total": len(contacts)})


@app.post("/api/contacts")
def api_create_contact():
    payload = request.get_json(silent=True) or {}
    if not str(payload.get("name", "")).strip():
        return jsonify({"ok": False, "message": "姓名不能为空"}), 400
    contact = db.create_contact(payload)
    return jsonify({"ok": True, "data": contact}), 201


@app.get("/api/contacts/<int:contact_id>")
def api_get_contact(contact_id: int):
    contact = db.get_contact(contact_id)
    if contact is None:
        return jsonify({"ok": False, "message": "联系人不存在"}), 404
    return jsonify({"ok": True, "data": contact})


@app.put("/api/contacts/<int:contact_id>")
def api_update_contact(contact_id: int):
    payload = request.get_json(silent=True) or {}
    if not str(payload.get("name", "")).strip():
        return jsonify({"ok": False, "message": "姓名不能为空"}), 400
    contact = db.update_contact(contact_id, payload)
    if contact is None:
        return jsonify({"ok": False, "message": "联系人不存在"}), 404
    return jsonify({"ok": True, "data": contact})


@app.delete("/api/contacts/<int:contact_id>")
def api_delete_contact(contact_id: int):
    if not db.delete_contact(contact_id):
        return jsonify({"ok": False, "message": "联系人不存在"}), 404
    return jsonify({"ok": True})


@app.patch("/api/contacts/<int:contact_id>/favorite")
def api_toggle_favorite(contact_id: int):
    payload = request.get_json(silent=True) or {}
    favorite = bool(payload.get("favorite"))
    contact = db.set_favorite(contact_id, favorite)
    if contact is None:
        return jsonify({"ok": False, "message": "联系人不存在"}), 404
    return jsonify({"ok": True, "data": contact})


@app.get("/api/groups")
def api_groups():
    return jsonify({"ok": True, "data": db.list_groups()})


@app.get("/api/stats")
def api_stats():
    return jsonify({"ok": True, "data": db.stats()})


@app.get("/api/export")
def api_export():
    contacts = db.list_contacts()
    buffer = excel_io.export_contacts(contacts)
    stamp = datetime.now().strftime("%Y%m%d")
    return send_file(
        buffer,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"通讯录_{stamp}.xlsx",
    )


@app.get("/api/template")
def api_template():
    buffer = excel_io.build_template()
    return send_file(
        buffer,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="通讯录导入模板.xlsx",
    )


@app.post("/api/import")
def api_import():
    file = request.files.get("file")
    if file is None or not file.filename:
        return jsonify({"ok": False, "message": "请选择要导入的 Excel 文件"}), 400

    filename = file.filename.lower()
    if filename.endswith(".xls"):
        return jsonify({"ok": False, "message": "暂不支持旧版 .xls，请在 Excel 中另存为 .xlsx 后再导入"}), 400
    if not filename.endswith(ALLOWED_SUFFIX):
        return jsonify({"ok": False, "message": "只支持 .xlsx 格式的表格"}), 400

    merge_same_name = request.form.get("merge", "1") in ("1", "true", "yes")

    try:
        # 统一读进 BytesIO 再解析：Python 3.10 的 SpooledTemporaryFile 没有实现
        # seekable()，而 zipfile / openpyxl 需要它，直接传 file.stream 会报
        # AttributeError（Ubuntu 22.04 自带的 python3.10 即如此）
        parsed, errors = excel_io.parse_workbook(io.BytesIO(file.read()))
    except Exception as exc:  # 文件损坏 / 不是 xlsx
        return jsonify({"ok": False, "message": f"文件解析失败：{exc}"}), 400

    imported = updated = 0
    for item in parsed:
        existing = db.find_by_name(item["name"]) if merge_same_name else []
        if existing:
            db.add_methods(existing[0]["id"], item["methods"])
            updated += 1
        else:
            db.create_contact(item)
            imported += 1

    return jsonify(
        {
            "ok": True,
            "data": {
                "imported": imported,
                "updated": updated,
                "skipped": len(errors),
                "errors": errors,
                "stats": db.stats(),
            },
        }
    )


if __name__ == "__main__":
    db.init_db()
    port = int(os.environ.get("PORT", "5000"))
    debug = os.environ.get("FLASK_DEBUG") == "1"
    print(f"* 通讯录已启动： http://127.0.0.1:{port}")
    app.run(host="0.0.0.0", port=port, debug=debug)
