#!/bin/zsh
cd "${0:A:h}"
if [[ ! -x .venv/bin/python ]]; then
  echo '首次使用：请先运行 python3 -m venv .venv，然后运行 .venv/bin/python -m pip install -r requirements.txt'
  read '?按回车关闭'
  exit 1
fi
.venv/bin/python reconcile.py
