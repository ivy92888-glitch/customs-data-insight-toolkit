@echo off
chcp 65001 >nul
title Customs Data Insight Toolkit
echo ============================================================
echo   Customs Data Insight Toolkit / 清关数据洞察工具包
echo ============================================================
echo.

REM ---- Step 1: 检查 Python ----
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [X] 没找到 python。请先装 Python 3.9+ 并勾选 Add to PATH。
    echo     https://www.python.org/downloads/
    pause
    exit /b 1
)
echo [OK] Python 已就绪

REM ---- Step 2: 安装依赖（只装一次）----
echo [..] 检查/安装依赖 ...
python -c "import pandas, matplotlib, seaborn, openpyxl, faker, docx" 2>nul
if %errorlevel% neq 0 (
    python -m pip install pandas numpy matplotlib seaborn openpyxl faker python-docx -q
)
echo [OK] 依赖已就绪

REM ---- Step 3: 运行分析（默认用 mock 数据，1 分钟内出结果）----
echo.
echo [..] 运行 mock 数据分析 ...
python supply_chain_analyzer.py --mock --rows 2000 --output output
if %errorlevel% neq 0 (
    echo [X] 分析失败，请把上面的错误贴给维护者。
    pause
    exit /b 1
)

echo [..] 生成 Word 报告 ...
python build_word_report.py --mock --rows 2000 --output output
if %errorlevel% neq 0 (
    echo [X] Word 报告生成失败。
    pause
    exit /b 1
)

REM ---- Done ----
echo.
echo ============================================================
echo   [DONE] 全部完成！
echo   打开 output\ 目录即可看到：
echo     - chart1_volumetric_loss.png    泡货 TOP10 + 渠道入仓
echo     - chart2_customer_profile.png  客户画像堆叠对比
echo     - chart3_data_quality.png       模糊词占比 + 人工成本
echo     - chart4_compliance_matrix.png  国家 × 报关方式矩阵
echo     - business_insight_report.md    Markdown 报告
echo     - supply_chain_report.docx      Word 报告
echo ============================================================
echo.
pause
