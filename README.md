# Customs Data Insight Toolkit

> A Python toolkit for multi-dimensional customs-clearance / supply-chain analytics,
> producing charts + Markdown + Word (docx) business-insight reports.

**Repo:** `customs-data-insight-toolkit` · GitHub org `ivy92888-glitch`

---

## 🎯 What this repo does

Given two worksheets in a single Excel workbook — **Loading List（装柜清单）**
and **Raw Data（原数据）** — the toolkit runs **4 business-critical dimensions
of analysis** and produces a management-ready report:

| Dimension | What we measure | Why management cares |
|-----------|-----------------|----------------------|
| **D1** Volumetric-loss & channel throughput | Cargo density (kg/CBM), light-bulk ratio, daily inbound qty by channel | "Bubble cargo" (ρ < 167) eats container space — fix it = direct profit lift |
| **D2** Customer profiling | Brand-awareness rate, material/usage cross by brand | Brand-owning customers = your high-NWI core; you either keep them or they churn |
| **D3** Data quality & manual-intervention cost | Vague-name scan (套装/配件/混装/其他…), cleaning rate, cost estimate | Hard ROI for building a "name → HS-code auto-matching engine" |
| **D4** Country × Customs-declaration matrix | Compliance matrix cross-tab + stacked bar chart | See at-a-glance which countries have compliance risk |

## 📁 Project layout

```
.
├── supply_chain_analyzer.py     # SupplyChainAnalyzer class + generate_mock_data()
├── build_word_report.py         # Builds a .docx report from analysis results
├── .gitignore
└── README.md
```

Dependencies:

```bash
python -m pip install pandas numpy matplotlib seaborn openpyxl python-docx faker
```

## 🚀 Quick start

### 1. Self-test with mock data (no real Excel needed)

```bash
python supply_chain_analyzer.py --mock --rows 2000 --output output_mock
python build_word_report.py  --mock --rows 2000 --output output_mock
```

### 2. Analyze your real workbook(s)

```bash
python supply_chain_analyzer.py --real --max-files 50 --output output_real
python build_word_report.py  --real --max-files 50 --output output_real
```

The real-data loader recursively scans `<用户真实数据目录>/<subfolder>/*已改*.xlsx`
and concatenates the two target sheets across all files found.

## 🧰 Design notes

- **PEP8 + OOP**: `SupplyChainAnalyzer` class, modular per-dimension methods
- **Chinese-language safe**: Matplotlib auto-picks a CJK font; Word builder
  writes `w:eastAsia` in every run so CJK text renders correctly
- **Resilient**: 3 bad files / table-shape mismatches are skipped gracefully;
  Density formula correctly uses **kg / CBM** (physical density), not its inverse
- **Self-contained**: `generate_mock_data(n_rows=2000, seed=42)` produces
  realistic data with Faker so any one can clone + run + reproduce the charts

## 🔒 Privacy / Data-security statement

1. All scripts have been sanitized — no former company name, channel code,
   brand, warehouse, or internal account remains.
2. The `.gitignore` blocks `*.xlsx`, `output*/`, `*.docx`, and any path
   containing `已上传云文档`. Push only Python source code.
3. If you fork and publish, **do not** commit your real Excel dataset.

## 📄 License

MIT — do whatever, attribution appreciated, no warranty.
