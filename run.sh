#!/usr/bin/env bash
# Customs Data Insight Toolkit — macOS / Linux 一键启动
set -e

echo "============================================================"
echo "  Customs Data Insight Toolkit / 清关数据洞察工具包"
echo "============================================================"
echo ""

# ---- Step 1: 检查 Python ----
if ! command -v python3 &>/dev/null; then
    echo "[X] 没找到 python3。请先装 Python 3.9+ (brew install python 或 https://www.python.org/downloads/)"
    exit 1
fi
PY=python3
echo "[OK] Python 已就绪"

# ---- Step 2: 依赖 ----
echo "[..] 检查/安装依赖 ..."
if ! $PY -c "import pandas, matplotlib, seaborn, openpyxl, faker, docx" 2>/dev/null; then
    $PY -m pip install pandas numpy matplotlib seaborn openpyxl faker python-docx -q
fi
echo "[OK] 依赖已就绪"

# ---- Step 3: 跑分析 ----
echo ""
echo "[..] 运行 mock 数据分析 ..."
$PY supply_chain_analyzer.py --mock --rows 2000 --output output

echo "[..] 生成 Word 报告 ..."
$PY build_word_report.py --mock --rows 2000 --output output

# ---- Done ----
echo ""
echo "============================================================"
echo "  [DONE] 全部完成！"
echo "  打开 output/ 目录即可看到图表、Markdown 和 Word 报告。"
echo "============================================================"
