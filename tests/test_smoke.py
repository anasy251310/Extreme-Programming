"""端到端冒烟测试：验证接口、Excel 导出导入往返一致性。

用法（先启动服务）：
    python tests/test_smoke.py
    BASE_URL=http://127.0.0.1:5000 python tests/test_smoke.py
"""

import io
import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request
import uuid

from openpyxl import Workbook, load_workbook

BASE_URL = os.environ.get("BASE_URL", "http://127.0.0.1:5000")
PASSED, FAILED = [], []

# 本机访问不走系统代理，避免校园网 / 公司代理把 127.0.0.1 也拦截掉
OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def check(name: str, condition: bool, detail: str = "") -> None:
    if condition:
        PASSED.append(name)
        print(f"  [PASS] {name}")
    else:
        FAILED.append(name)
        print(f"  [FAIL] {name} {detail}")


def request(method: str, path: str, payload=None, raw: bytes | None = None, headers=None):
    url = BASE_URL + urllib.parse.quote(path, safe="/?=&%:")
    data = None
    hdrs = dict(headers or {})
    if raw is not None:
        data = raw
    elif payload is not None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        hdrs["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=data, headers=hdrs, method=method)
    with OPENER.open(req, timeout=20) as resp:
        body = resp.read()
        if resp.headers.get("Content-Type", "").startswith("application/json"):
            return resp.status, json.loads(body.decode("utf-8"))
        return resp.status, body


def multipart(field_name: str, filename: str, content: bytes, extra: dict | None = None) -> tuple[bytes, dict]:
    boundary = "----smoke" + uuid.uuid4().hex
    lines = []
    for key, value in (extra or {}).items():
        lines.append(f"--{boundary}\r\nContent-Disposition: form-data; name=\"{key}\"\r\n\r\n{value}\r\n".encode())
    lines.append(
        f"--{boundary}\r\nContent-Disposition: form-data; name=\"{field_name}\"; filename=\"{filename}\"\r\n"
        f"Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n".encode()
    )
    lines.append(content)
    lines.append(f"\r\n--{boundary}--\r\n".encode())
    return b"".join(lines), {"Content-Type": f"multipart/form-data; boundary={boundary}"}


def main() -> int:
    print(f"== 目标服务：{BASE_URL}")

    print("\n[1] 基础接口")
    status, stats = request("GET", "/api/stats")
    check("GET /api/stats 返回 200", status == 200)
    total_before = stats["data"]["total"]
    check("联系人数量大于 0", total_before > 0, f"total={total_before}")

    print("\n[2] 新建联系人 + 多联系方式 + 收藏")
    payload = {
        "name": "测试用联系人" + uuid.uuid4().hex[:4],
        "group": "自动化测试",
        "company": "测试公司",
        "note": "由冒烟测试创建",
        "favorite": False,
        "methods": [
            {"type": "phone", "label": "手机", "value": "13800001234"},
            {"type": "email", "label": "工作邮箱", "value": "qa@example.com"},
            {"type": "wechat", "label": "", "value": "qa_wx"},
            {"type": "address", "label": "公司", "value": "福州市闽侯县"},
        ],
    }
    status, created = request("POST", "/api/contacts", payload)
    check("POST /api/contacts 201", status == 201)
    contact_id = created["data"]["id"]
    check("联系方式全部写入", len(created["data"]["methods"]) == 4, str(created["data"]["methods"]))

    status, fav = request("PATCH", f"/api/contacts/{contact_id}/favorite", {"favorite": True})
    check("PATCH 收藏成功", status == 200 and fav["data"]["favorite"] is True)

    status, listed = request("GET", "/api/contacts?favorite=1")
    check("收藏筛选包含新联系人", any(c["id"] == contact_id for c in listed["data"]))

    status, listed = request("GET", "/api/contacts?keyword=qa@example.com")
    check("按联系方式搜索命中", any(c["id"] == contact_id for c in listed["data"]))

    print("\n[3] 编辑联系人（PUT 为全量更新，带上收藏状态）")
    payload["name"] += "改"
    payload["favorite"] = True
    payload["methods"] = payload["methods"][:3]
    status, updated = request("PUT", f"/api/contacts/{contact_id}", payload)
    check("PUT 更新成功", status == 200 and updated["data"]["name"].endswith("改"))
    check("联系方式数量被更新", len(updated["data"]["methods"]) == 3)

    print("\n[4] 导出 Excel")
    status, blob = request("GET", "/api/export")
    check("GET /api/export 返回 200", status == 200)
    check("导出内容非空", len(blob) > 2000, f"size={len(blob)}")
    wb = load_workbook(io.BytesIO(blob))
    ws = wb["通讯录"]
    headers = [c.value for c in ws[1]]
    check("工作表名为「通讯录」", "通讯录" in wb.sheetnames)
    check("包含姓名 / 分组 / 收藏列", all(h in headers for h in ("姓名", "分组", "收藏")))
    check("联系方式按类型 + 内容分列", "联系方式1类型" in headers and "联系方式1内容" in headers)
    check("导出行数 = 联系人数 + 1", ws.max_row == total_before + 2, f"rows={ws.max_row}")
    exported_contact = next(
        (row for row in ws.iter_rows(min_row=2, values_only=True) if row[0] == updated["data"]["name"]),
        None,
    )
    check("导出包含新联系人", exported_contact is not None)
    if exported_contact:
        check("导出收藏列为「是」", exported_contact[2] == "是", str(exported_contact))
        check("导出联系方式内容正确", exported_contact[5] == "13800001234", str(exported_contact[5]))

    print("\n[5] 导入刚导出的文件（同名合并，不应产生重复）")
    body, headers_mp = multipart("file", "export.xlsx", blob, {"merge": "1"})
    status, result = request("POST", "/api/import", raw=body, headers=headers_mp)
    check("POST /api/import 200", status == 200, str(result)[:200])
    check("合并模式没有新增重复联系人", result["data"]["imported"] == 0, str(result["data"])[:200])
    check("合并数量等于导出行数", result["data"]["updated"] == total_before + 1, str(result["data"])[:200])
    status, stats2 = request("GET", "/api/stats")
    check("导入后联系人总数不变", stats2["data"]["total"] == total_before + 1, str(stats2["data"]))

    print("\n[6] 导入「手机 / 邮箱 / 微信 各占一列」的自定义表格")
    wb2 = Workbook()
    sheet = wb2.active
    sheet.title = "通讯录"
    sheet.append(["姓名", "分组", "手机号", "邮箱", "微信", "QQ", "地址", "备注"])
    sheet.append(["外部导入甲", "批量导入", "13911112222", "jia@example.com", "jia_wx", "", "厦门市思明区", "来自简易表头"])
    sheet.append(["外部导入乙", "批量导入", "13933334444；13955556666", "", "", "778899", "", ""])
    sheet.append(["", "", "", "", "", "", "", "姓名为空应被跳过"])
    buf = io.BytesIO()
    wb2.save(buf)
    body, headers_mp = multipart("file", "custom.xlsx", buf.getvalue(), {"merge": "0"})
    status, result = request("POST", "/api/import", raw=body, headers=headers_mp)
    check("自定义表头导入成功", status == 200 and result["data"]["imported"] == 2, str(result["data"])[:300])
    check("空姓名行被跳过并记录", result["data"]["skipped"] >= 1, str(result["data"]["errors"]))

    status, listed = request("GET", "/api/contacts?group=批量导入")
    target = next((c for c in listed["data"] if c["name"] == "外部导入乙"), None)
    phones = [m for m in (target or {}).get("methods", []) if m["type"] == "phone"]
    check("多值单元格被拆成两条手机号", len(phones) == 2, str(target))

    print("\n[7] 下载模板")
    status, tpl = request("GET", "/api/template")
    check("GET /api/template 返回 200", status == 200 and len(tpl) > 2000)

    print("\n[8] 删除测试数据")
    for item in request("GET", "/api/contacts")[1]["data"]:
        if item["group"] == "批量导入" or item["name"].startswith("测试用联系人"):
            request("DELETE", f"/api/contacts/{item['id']}")

    try:
        request("GET", f"/api/contacts/{contact_id}")
        check("删除后查询返回 404", False, "接口仍能查到已删除的联系人")
    except urllib.error.HTTPError as exc:
        check("删除后查询返回 404", exc.code == 404, f"code={exc.code}")

    status, stats3 = request("GET", "/api/stats")
    check("清理后回到初始数量", stats3["data"]["total"] == total_before, str(stats3["data"]))

    print(f"\n== 结果：{len(PASSED)} 通过 / {len(FAILED)} 失败")
    if FAILED:
        print("失败项：" + "、".join(FAILED))
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
