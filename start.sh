#!/usr/bin/env bash
# 一键安装依赖并启动通讯录
set -e
cd "$(dirname "$0")"

if [ ! -d ".venv" ]; then
  echo "[1/3] 创建虚拟环境..."
  python3 -m venv .venv
fi

echo "[2/3] 检查依赖..."
# shellcheck disable=SC1091
source .venv/bin/activate
pip install -q --disable-pip-version-check -r requirements.txt

echo "[3/3] 启动通讯录： http://127.0.0.1:5000"
python app.py
