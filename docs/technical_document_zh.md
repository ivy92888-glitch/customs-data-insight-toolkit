# 清关数据洞察工具包 · 技术文档

> 本文档基于源码 `supply_chain_analyzer.py`（1134 行）与 `build_word_report.py`（664 行）抽取编写，所有接口签名、数据结构、字段名均为源码事实。

---

## 1. 项目总览

### 1.1 这是什么

一个面向清关公司 / 供应链团队的 Python 数据分析工具包。输入是一个文件夹里几十份 Excel 清关资料（每份都有"装柜清单"和"原数据"两张表），输出是 **4 张数据图表 + 一份 Markdown 洞察报告 + 一份 Word (.docx) 业务报告**。

### 1.2 核心目标

不是纯学术展示——而是要向管理层证明两件事：

1. **为什么要立即做内部流程优化**（泡货吞噬集装箱空间、无单证合规风险）
2. **为什么要投入资源开发自动化工具**（模糊品名导致的人工 HS 编码匹配成本可量化）

### 1.3 设计原则

- **零依赖部署**：`Faker` 可选；`python-docx` 可选（只生成 Markdown 也能跑）
- **降级友好**：真实数据目录不存在时自动降级到 mock 模式
- **无 GUI**：`matplotlib.use("Agg")` 头less 模式，无显示器的服务器也能跑
- **中文优先**：Matplotlib 和 python-docx 都有多重 CJK 字体 fallback
- **安全默认**：所有硬编码路径脱敏（`E:\已上传云文档` → `<用户真实数据目录>`）

### 1.4 技术栈

| 组件 | 版本要求 | 用途 |
|------|---------|------|
| Python | 3.9+ | 运行时 |
| pandas | ≥ 2.0 | 数据清洗、groupby、crosstab |
| numpy | ≥ 1.24 | 数值计算（密度、统计量）|
| matplotlib | ≥ 3.7 | 图表绘制（4 张）|
| seaborn | ≥ 0.12 | 调色板 & 绘图风格 |
| openpyxl | ≥ 3.1 | 读取 .xlsx（真实数据 loader）|
| python-docx | ≥ 1.1 | 生成 Word 报告 |
| Faker | ≥ 18.0（可选）| Mock 数据的中国地名/地址生成 |

一行装完：

```bash
python -m pip install pandas numpy matplotlib seaborn openpyxl python-docx faker
```

---

## 2. 架构设计

### 2.1 数据流图

```
┌─────────────────────────────────────────────────────────────────┐
│                        输入层                                     │
│                                                                   │
│  真实模式           Mock 模式                                     │
│  ┌─────────────────┐  ┌────────────────────┐                     │
│  │ 扫描 <用户真实   │  │ generate_mock_data │                     │
│  │ 数据目录>/**/    │  │ (Faker + 业务枚举)  │                     │
│  │ *已改*.xlsx      │  └─────────┬──────────┘                     │
│  └────────┬────────┘            │                                │
│           │                      │                                │
│  ┌────────▼────────┐            │                                │
│  │ load_real_data  │            │                                │
│  │ glob + header=9  │            │                                │
│  │ 两表合并 concat  │            │                                │
│  └────────┬────────┘            │                                │
│           └──────────┬───────────┘                                │
└──────────────────────┼───────────────────────────────────────────┘
                       │
                       ▼
┌─────────────────────────────────────────────────────────────────┐
│  {"装柜清单": DataFrame, "原数据": DataFrame}                     │
│                        ↓                                         │
│                   SupplyChainAnalyzer(data)                       │
│                                                                  │
│  ┌─────────────────────────────────────────────────────────┐     │
│  │  维度 1  密度计算 + 泡货标记 + TOP10 + 渠道均值         │     │
│  │  维度 2  品牌赋能率 + 材质交叉 + 用途交叉                │     │
│  │  维度 3  模糊词扫描 + 清洗率 + 人工成本估算              │     │
│  │  维度 4  运单 merge + 国家×报关方式 crosstab             │     │
│  └─────────────────────────────────────────────────────────┘     │
│                        │                                         │
│               self.results (dict)                                 │
│               4 张 PNG 图表                                       │
│               business_insight_report.md                         │
└──────────────────────┬──────────────────────────────────────────┘
                       │
          ┌────────────┴────────────┐
          ▼                         ▼
   Markdown 报告              Word 报告 (.docx)
   (analyzer 内置)            (build_word_report.py)
                               python-docx + CJK XML 注入
```

### 2.2 模块划分

| 模块 | 文件 | 行数 | 职责 |
|------|------|------|------|
| Mock 生成器 | `supply_chain_analyzer.py` L103-L297 | 195 | 生成两张 DataFrame（含共享运单号池）|
| 真实数据 Loader | `supply_chain_analyzer.py` L304-L387 | 84 | glob 扫描 + 表结构解析 + 数值类型转换 |
| 核心分析类 | `supply_chain_analyzer.py` L394-L1065 | 672 | 4 维度方法 + run_all() + Markdown 报告生成 |
| CLI 入口 | `supply_chain_analyzer.py` L1072-L1134 | 63 | argparse + auto-detect mock/real |
| Word 构建 helpers | `build_word_report.py` L51-L333 | 283 | CJK 字体系统 + 6 个组件函数 |
| Word 报告主函数 | `build_word_report.py` L340-L616 | 277 | 封面 + 4 维度 + 总结 |

### 2.3 调用关系

```
main()
 ├─ [--mock]   generate_mock_data(n_rows, seed)
 ├─ [--real]   load_real_data(root_dir, max_files)
 └─ SupplyChainAnalyzer(data, output_dir, source_label)
      ├─ run_all()
      │    ├─ analyze_dim1_volumetric()   → chart1_volumetric_loss.png
      │    ├─ analyze_dim2_customer()     → chart2_customer_profile.png
      │    ├─ analyze_dim3_data_quality() → chart3_data_quality.png
      │    └─ analyze_dim4_compliance()   → chart4_compliance_matrix.png
      └─ generate_report()                → business_insight_report.md

build_word_report.py main()
 ├─ load_real_data / generate_mock_data
 ├─ SupplyChainAnalyzer(...).run_all()   （Word 脚本独立跑一次分析）
 └─ build_word_report(analyzer, output_dir) → supply_chain_report.docx
```

---

## 3. 数据结构

### 3.1 Mock 数据业务枚举（源码事实）

| 枚举 | 值 | 用途 |
|------|---|------|
| `WAREHOUSE_CODES` | DTM1, KSF7, WRO5, DUS4, BER5, HAM3, FRA2, ROT3, LAX8, JFK5, LHR2, CDG3, MAD2, BAR4 | 仓库代码（装柜清单 D 列）|
| `CHANNEL_ARRANGEMENTS` | 渠道-01 ~ 渠道-07 | 渠道安排（装柜清单 M 列）|
| `DEST_COUNTRIES` | 德国/英国/法国/荷兰/波兰/西班牙/意大利/美国/加拿大/澳大利亚/日本/韩国 | 收件国家（原数据 Q 列）|
| `BRAND_POOL` | GenericBrand-A ~ GenericBrand-J | 脱敏后品牌池 |
| `CUSTOMS_DECLARATION` | 有单证 / 无单证 | 报关方式（装柜清单 N 列）|
| `FUZZY_KEYWORDS` | 套装/配件/混装/其他/杂项/零件 | 维度 3 模糊词扫描列表 |

### 3.2 共享运单号池（关键设计）

```python
n_waybills = max(200, n_rows // 3)
waybill_ids = [f"99926{random.randint(1000000000, 9999999999)}"
               for _ in range(n_waybills)]
```

**为什么共享？** 业务上一个运单号对应装柜清单的一行，也对应原数据的多行（每件商品一行）。维度 4 要把两表 merge 起来做国家 × 报关方式交叉——没有共享运单号就 merge 不上。

### 3.3 真实数据表结构（源码事实）

**表 1：装柜清单**（Excel 第 9 行为表头）

| 列 | 字段名 | 类型 | 用途 |
|----|--------|------|------|
| D | 仓库代码 | str | 维度 1 分组 |
| E | 地址 | str | 留空 |
| F | 件数 | int | 维度 1 sum() |
| G | 品名 | str | 维度 1 groupby（TOP10）|
| H | 总KG | float | 维度 1 密度分母 |
| I | 总CBM | float | 维度 1 密度分子 |
| J | 入仓时间 | datetime | 维度 1 日期截断 → 渠道日均 |
| M | 渠道安排 | str | 维度 1 分组 |
| N | 报关方式 | str | 维度 4 分组 |
| — | 运单号 | str | 维度 4 merge key |

**表 2：原数据**（Excel 第 0 行为表头）

| 列 | 字段名 | 类型 | 用途 |
|----|--------|------|------|
| B | 中文品名 | str | 维度 3 模糊词扫描 |
| G | 箱单件数 | int | 留空 |
| H | 收货实重 | float | 留空 |
| I | 品牌 | str | 维度 2 有/无品牌标记 |
| K | 仓库编码 | str | 留空 |
| L | 材质 | str | 维度 2 交叉 |
| M | 用途 | str | 维度 2 交叉 |
| Q | 收件国家 | str | 维度 4 国家来源 |
| — | 运单号码 | str | 维度 4 merge key |

### 3.4 密度公式（维度 1 核心）

```python
df["密度"] = df["总KG"].replace(0, np.nan) / df["总CBM"].replace(0, np.nan)
df["是否泡货"] = df["密度"] < LIGHTER_THAN_DENSITY_THRESHOLD   # 167.0
```

**为什么是 KG/CBM？** 物流行业的体积重公式是 `体积重 = CBM × 167 kg/CBM`。当实际重量 KG < 体积重时，计费重取体积重——这就是"泡货"的定义。等价于密度 KG/CBM < 167。密度越低越泡。

---

## 4. 核心类与 API

### 4.1 `SupplyChainAnalyzer` 类

**构造函数**

```python
SupplyChainAnalyzer(
    data: Dict[str, pd.DataFrame],   # {"装柜清单": df, "原数据": df}
    output_dir: str = "output",
    source_label: str = "mock"
)
```

**实例属性**

| 属性 | 类型 | 说明 |
|------|------|------|
| `df_loading` | DataFrame | 装柜清单（copy，安全）|
| `df_raw` | DataFrame | 原数据（copy，安全）|
| `output_dir` | Path | 图表和报告的输出目录 |
| `source_label` | str | `"mock"` / `"real"` |
| `results` | dict | 4 维度的结构化分析结果（供 Word 报告读取）|

**类常量**

```python
SupplyChainAnalyzer.FUZZY_KEYWORDS = ["套装", "配件", "混装", "其他", "杂项", "零件"]
```

**方法签名**

| 方法 | 输入 | 输出 | 副作用 |
|------|------|------|--------|
| `analyze_dim1_volumetric()` | — | — | `results["dim1"]` + 保存 chart1 PNG |
| `analyze_dim2_customer()` | — | — | `results["dim2"]` + 保存 chart2 PNG |
| `analyze_dim3_data_quality()` | — | — | `results["dim3"]` + 保存 chart3 PNG |
| `analyze_dim4_compliance()` | — | — | `results["dim4"]` + 保存 chart4 PNG |
| `run_all()` | — | — | 依次调用以上 4 个方法 |
| `generate_report()` | — | `Path` | 生成 Markdown 报告并返回路径 |

### 4.2 `results` 字典结构（供 Word 报告消费）

```python
{
    "dim1": {
        "bubble_ratio_pct": float,           # 泡货比例
        "top10_lightest_products": Dict[str, float],  # {品名: 平均密度}
        "channel_daily_avg_qty": Dict[str, float],    # {渠道: 日均件数}
        "df_density": DataFrame,             # 带密度列的 DataFrame
    },
    "dim2": {
        "brand_rate_pct": float,             # 品牌赋能率
        "material_cross_pct": DataFrame,     # 材质 × 有品牌（行归一化 %）
        "usage_cross_pct": DataFrame,        # 用途 × 有品牌（行归一化 %）
        "n_brand": int,
        "n_no_brand": int,
    },
    "dim3": {
        "total_rows": int,
        "needs_cleaning_rows": int,
        "cleaning_rate_pct": float,
        "per_keyword_pct": Dict[str, float], # {关键词: 占比 %}
        "sample_dirty_names": List[str],     # 前 15 个模糊品名（供抽查）
    },
    "dim4": {
        "cross_pct": DataFrame,              # 国家 × 报关方式（行归一化 %）
        "cross_count": DataFrame,            # 国家 × 报关方式（原始计数）
        "total_by_country": Dict[str, int],  # {国家: 合计}
    },
}
```

### 4.3 工具函数签名

| 函数 | 文件 | 输入 | 输出 |
|------|------|------|------|
| `generate_mock_data(n_rows=2000, seed=42)` | 主脚本 L103 | 行数、种子 | `{"装柜清单": DataFrame, "原数据": DataFrame}` |
| `load_real_data(root_dir, max_files=None)` | 主脚本 L304 | 根目录、文件上限 | 同上 |
| `build_word_report(analyzer, output_dir)` | Word 脚本 L340 | analyzer 实例、输出目录 | `Path`（.docx 路径）|

### 4.4 Word 构建器组件（helper 函数）

| 函数 | 作用 | 关键技术 |
|------|------|---------|
| `_set_run_cjk_font(run, font_name, ascii_font)` | 为 python-docx 的 run 同时设置 ASCII 和 CJK 字体 | 直接写 XML：`w:eastAsia` / `w:ascii` / `w:hAnsi` |
| `add_banner_quote(doc, text, bg_color, text_color)` | 管理层建议高亮条（表格模拟） | 单单元格表格 + `parse_xml(w:shd)` + 去边框 |
| `add_table(doc, headers, rows, col_widths_cm, header_color)` | 带蓝底白字表头 + 斑马纹通用表格 | XML 写 `w:tblBorders` 和 `w:shd` |
| `add_image_with_caption(doc, path, caption, width_inches)` | 居中插图 + 斜体灰色图题 | `run.add_picture()` + `ImageRun` |

---

## 5. CLI 用法

### 5.1 主分析器

```bash
python supply_chain_analyzer.py                    # 自动检测（真实目录存在则真实，否则 mock）
python supply_chain_analyzer.py --mock --rows 5000 # 强制 mock，加大数据量
python supply_chain_analyzer.py --real --max-files 30 --output output_real
```

**参数**

| 参数 | 默认 | 说明 |
|------|------|------|
| `--mock` | - | 强制模拟数据 |
| `--real` | - | 强制真实数据 |
| `--seed` | `42` | Mock 随机种子 |
| `--rows` | `2000` | Mock 每个表的行数 |
| `--max-files` | `None` | 真实模式最多读取文件数 |
| `--output` | `output` | 输出目录 |

### 5.2 Word 报告

```bash
python build_word_report.py --mock --output output_mock
python build_word_report.py --real --max-files 30 --output output_real
```

Word 脚本内部会自己跑一遍 `SupplyChainAnalyzer.run_all()`，不需要先跑主分析器。

### 5.3 run.bat / run.sh

一键启动脚本，自动装依赖 → 跑 mock → 出结果。用户零配置。

---

## 6. 设计决策说明

### 6.1 Matplotlib 中文字体：四重 fallback

```python
_CN_FONT_CANDIDATES = [
    "Microsoft YaHei",   # Windows 10+ 通用
    "SimHei", "SimSun", "KaiTi",
    "Arial Unicode MS",  # Mac
    "DejaVu Sans",       # 兜底（会显示英文，中文可能方框）
]
```

**为什么？** Windows 机器上字体安装情况千差万别。先扫 `matplotlib.font_manager` 找到本机已装的第一个候选，找不到就 warn。

### 6.2 python-docx 中文渲染：必须写 XML

python-docx 的 `run.font.name` 只能设 ASCII 字体。CJK 字体必须手动写 `w:eastAsia`：

```python
rFonts.set(qn("w:eastAsia"), font_name)
rFonts.set(qn("w:ascii"), "Arial")
rFonts.set(qn("w:hAnsi"), "Arial")
```

**不写会怎么样？** Word 打开中文全是方框 □□□。

### 6.3 真实数据"装柜清单"表：header=9

```python
pd.read_excel(fpath, sheet_name="装柜清单", header=9, engine="openpyxl")
```

**为什么是 9？** 用户提供的 Excel 模板前 9 行是公司抬头、Logo、客户信息等非表头内容。第 10 行（0-indexed 9）才是真正的列名。跳过这个参数读出来全是 None。

### 6.4 维度 4 的运单 merge：先 groupby("运单号码") 再取 first

```python
waybill_country = (
    df_raw.dropna(subset=["运单号码"])
          .groupby("运单号码", dropna=False)["收件国家"]
          .first()
)
df_load = df_load.merge(waybill_country.rename("收件国家"),
                       left_on="运单号", right_index=True, how="left")
```

**为什么？** 原数据一张运单对应多行（每件商品一行），每行都有同一个收件国家。先聚合避免 merge 时行数膨胀。用 `.first()` 是因为 groupby 内部顺序保留，且同一运单的收件国家一定一致。

### 6.5 Faker 可选

```python
try:
    from faker import Faker
    HAS_FAKER = True
except ImportError:
    HAS_FAKER = False
```

Mock 数据里地址生成有 Faker 就用 `fake.zh_CN.address()`，没有就退化成 `Street X, CityY, DE` 这种占位符。不影响任何分析逻辑。

### 6.6 密度公式为什么用 KG/CBM（物理密度）

而不是 CBM/KG（倒密度）。

物流行业的体积重公式是 `体积重 = CBM × 167`。物理密度 `KG/CBM < 167` 与"实际重量 < 体积重"是数学等价的——用物理密度更直观（密度越低越泡，密度越高越重货）。

---

## 7. 扩展指南

### 7.1 添加新分析维度

继承 `SupplyChainAnalyzer`，加一个方法：

```python
class SupplyChainAnalyzerV2(SupplyChainAnalyzer):
    def analyze_dim5_new(self) -> None:
        """维度 5：XXXXX"""
        df = self.df_loading.copy()
        # ... 计算 ...
        self.results["dim5"] = {
            "some_metric": value,
        }
        # 绘图 ...
```

然后在 `run_all()` 里加一行调用。

### 7.2 接入不同 Excel 结构

改 `load_real_data()` 的两个 header 参数：

```python
df_ld = pd.read_excel(fpath, sheet_name="装柜清单", header=9, engine="openpyxl")
df_rd = pd.read_excel(fpath, sheet_name="原数据",   header=0, engine="openpyxl")
```

如果表头位置变了（比如用户新模板把表头挪到了第 5 行），改 `header` 数字即可。

### 7.3 接入不同业务阈值

改模块顶部的常量：

```python
LIGHTER_THAN_DENSITY_THRESHOLD = 167.0   # 改这个能适配不同物流商
```

### 7.4 切换脱敏程度

Mock 数据里的品牌池 `GenericBrand-A~J`、渠道池 `渠道-01~07` 可以改成真实业务枚举来做真实分析。**前提是**这些枚举值不涉及任何第三方知识产权（建议保持脱敏形式用于开源）。

---

## 8. 已知限制

| 限制 | 原因 | 可能改进 |
|------|------|---------|
| Mock 数据的"模糊品名"比例固定 30% | 业务假设 | 做成参数 |
| 真实数据扫描硬编码 `*已改*.xlsx` | 适配用户文件命名约定 | 加 `--glob` 参数 |
| Word 报告的管理层建议文本硬编码在函数里 | 快速交付 | 抽到 JSON/YAML 配置 |
| 只出 PNG | Matplotlib 默认 | 加 PDF / SVG 导出 |
| 维度 4 merge key 只认"运单号"和"运单号码" | 用户数据字段名 | 加字段名映射表 |
| Mock 数据没有真实时间序列波动 | 均匀随机 | 引入月度季节性 |
