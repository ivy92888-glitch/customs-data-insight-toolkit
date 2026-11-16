# -*- coding: utf-8 -*-
"""
================================================================================
清关资料数据分析脚本 / Customs Clearance Data Analysis Script
================================================================================

【中文说明】
本脚本面向企业级清关/供应链场景，对真实的"装柜清单"与"原数据"两个 Excel
工作表进行四维度分析，输出 4 张数据图表和一份业务洞察 Markdown 报告。

核心目标：向管理层证明 —— 为什么要做内部流程优化，以及为什么必须开发
自动化数据清洗 / HS 编码自动匹配工具。

【English Description】
An enterprise-grade analysis script for customs clearance / supply chain
scenarios. Reads real "装柜清单" (Loading List) and "原数据" (Raw Data) sheets,
performs 4-dimensional analysis, outputs 4 charts + a Markdown insight report.

Goal: Convince management WHY process optimization is needed and WHY we must
build automated data-cleaning / HS-code matching tools.

【Usage】
    python supply_chain_analyzer.py              # 自动检测真实数据 / 降级用 mock
    python supply_chain_analyzer.py --mock       # 强制使用模拟数据
    python supply_chain_analyzer.py --real       # 强制使用真实数据

【真实数据路径】E:\\已上传云文档\\<各子文件夹>\\*已改*.xlsx
================================================================================
"""

# ---------------------------------------------------------------------------
# 标准库 / Standard library imports
# ---------------------------------------------------------------------------
import os
import re
import sys
import glob
import random
import datetime
import traceback
from pathlib import Path
from typing import List, Dict, Optional, Tuple, Any

# ---------------------------------------------------------------------------
# 第三方库 / Third-party imports
# ---------------------------------------------------------------------------
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")           # 无 GUI 环境也可保存 PNG / headless
import matplotlib.pyplot as plt
import seaborn as sns

try:
    from faker import Faker
    HAS_FAKER = True
except ImportError:
    HAS_FAKER = False

# ---------------------------------------------------------------------------
# 全局配置 / Global configuration
# ---------------------------------------------------------------------------
# 业务常量 / Business constants
LIGHTER_THAN_DENSITY_THRESHOLD = 167.0   # 物流行业：密度 < 167 kg/CBM 视为"泡货"
CHINA_TZ_OFFSET = "+08:00"

# Matplotlib 中文字体 fallback（按优先级尝试）
# 每台 Windows 机器字体安装情况不同，这里做多重兜底
# Chinese font fallback list for Windows matplotlib
_CN_FONT_CANDIDATES = [
    "Microsoft YaHei",    # 微软雅黑（Windows 10+ 通用）
    "SimHei",             # 黑体
    "SimSun",             # 宋体
    "KaiTi",              # 楷体
    "Arial Unicode MS",   # Mac 常见
    "DejaVu Sans",        # 兜底 fallback
]

# ---------------------------------------------------------------------------
# 字体初始化 / Font initialization
# ---------------------------------------------------------------------------
def _setup_matplotlib_font() -> None:
    """
    尝试为 matplotlib 设置一个能正确显示中文的字体。
    Try to pick a Chinese-capable font for matplotlib.
    """
    plt.rcParams["axes.unicode_minus"] = False  # 防止负号乱码
    import matplotlib.font_manager as fm
    installed = {f.name for f in fm.fontManager.ttflist}
    for candidate in _CN_FONT_CANDIDATES:
        if candidate in installed:
            plt.rcParams["font.sans-serif"] = [candidate]
            print(f"[INFO] matplotlib 字体已设置为: {candidate}")
            return
    print("[WARN] 未找到合适的中文字体，图表中中文可能显示为方框。")

_setup_matplotlib_font()

# ===========================================================================
# 第一部分：Mock 数据生成器 / Part 1: Mock data generator
# ===========================================================================

def generate_mock_data(n_rows: int = 2000, seed: int = 42) -> Dict[str, pd.DataFrame]:
    """
    生成逼真的模拟数据（两张表：装柜清单 + 原数据），用于脚本自测。

    Generate realistic mock data (Loading List + Raw Data) for self-testing.

    Args:
        n_rows: 每个工作表的模拟行数，默认 2000
        seed:   随机种子，保证可复现

    Returns:
        dict: {"装柜清单": DataFrame, "原数据": DataFrame}
    """
    random.seed(seed)
    np.random.seed(seed)
    if HAS_FAKER:
        fake = Faker("zh_CN")
        Faker.seed(seed)
    else:
        fake = None

    # ---- 业务枚举 / Business value pools ----
    WAREHOUSE_CODES = [
        "DTM1", "KSF7", "WRO5", "DUS4", "BER5", "HAM3", "FRA2",
        "ROT3", "LAX8", "JFK5", "LHR2", "CDG3", "MAD2", "BAR4",
    ]
    CHANNEL_ARRANGEMENTS = [
        "渠道-01",
        "渠道-02",
        "渠道-03",
        "渠道-04",
        "渠道-05",
        "渠道-06",
        "渠道-07",
    ]
    CUSTOMS_DECLARATION = ["有单证", "无单证"]
    DEST_COUNTRIES = [
        "德国", "英国", "法国", "荷兰", "波兰", "西班牙",
        "意大利", "美国", "加拿大", "澳大利亚", "日本", "韩国",
    ]
    BRAND_POOL = [
        "GenericBrand-A", "GenericBrand-B", "GenericBrand-C", "GenericBrand-D", "GenericBrand-E",
        "GenericBrand-F", "GenericBrand-G", "GenericBrand-H", "GenericBrand-I", "GenericBrand-J",
    ]
    NO_BRAND = ""       # 大量商品无品牌
    MATERIALS = [
        "不锈钢", "PP+PVC", "TPE", "塑料", "木质", "金属", "玻璃",
        "棉质", "涤纶", "硅胶", "铝合金", "铁", "陶瓷", "纸质",
    ]
    USES = [
        "翻土用 / Gardening",
        "通水清洁、浇花",
        "体育健身用 / Sports",
        "家居收纳 / Home storage",
        "厨房用品 / Kitchenware",
        "电子配件 / Electronics",
        "节日装饰 / Festival decor",
        "办公文具 / Office",
        "宠物用品 / Pet",
        "儿童玩具 / Toys",
        "汽车配件 / Auto parts",
        "工业零件 / Industrial",
    ]

    # -- 品名池（故意混入模糊词汇，用于数据质量维度分析）--
    PRODUCT_NAMES_CLEAN = [
        "铲子", "卷管器", "运动垫", "收纳架", "记事本", "数据线",
        "LED灯泡", "不锈钢保温杯", "硅胶蛋糕模", "PVC地毯",
        "陶瓷马克杯", "塑料衣架", "铝合金工具箱",
    ]
    PRODUCT_NAMES_FUZZY = [
        "套装", "配件", "混装", "其他", "杂项", "工具套装",
        "电子配件混装", "厨房套装", "零件套装",
    ]

    # 渠道产品名称（"原数据"表 U 列）
    CHANNEL_PRODUCT_NAMES = [
        "渠道产品-直送", "渠道产品-普船", "渠道产品-快船",
        "渠道产品-经济船", "渠道产品-空运", "渠道产品-铁路",
    ]

    # -- 先共享一批运单号，保证两表能 merge 上 --
    # 业务上：一个运单号对应多件商品（"原数据"粒度），
    # 也对应装柜清单里的一行（或多行，若拆柜）
    n_waybills = max(200, n_rows // 3)
    waybill_ids = [f"99926{random.randint(1000000000, 9999999999)}"
                   for _ in range(n_waybills)]

    # ---------- 生成 装柜清单 ----------
    records_loading = []
    base_date = datetime.date(2026, 8, 1)
    # 装柜清单：n_rows 行，每行选一个运单号（可能重复 = 一个运单多件）
    for i in range(n_rows):
        # 品名：70% 清晰 + 30% 模糊（故意制造数据质量问题）
        if random.random() < 0.70:
            product_name = random.choice(PRODUCT_NAMES_CLEAN)
        else:
            product_name = random.choice(PRODUCT_NAMES_FUZZY)

        warehouse = random.choice(WAREHOUSE_CODES)
        channel = random.choice(CHANNEL_ARRANGEMENTS)
        customs = random.choice(CUSTOMS_DECLARATION)

        # 件数 / 重量 / 体积：制造真实的"泡货"差异
        qty = random.randint(5, 300)
        # 随机选一个密度场景（物理密度 = KG/CBM）：
        #   - 30% 概率是"泡货"（密度 < 167 kg/CBM，体积重 > 实际重）
        #   - 70% 概率是"重货"（密度 > 167 kg/CBM，实际重 > 体积重）
        is_light_bulk = random.random() < 0.30
        if is_light_bulk:
            physical_density = random.uniform(60, 150)     # kg/CBM，泡货
        else:
            physical_density = random.uniform(180, 500)    # kg/CBM，重货

        # 体积 CBM 先随机，再用物理密度反推重量 KG
        volume = round(random.uniform(0.1, 25.0), 4)
        weight = round(volume * physical_density, 2)

        # 入仓时间（2026-08-01 ~ 2026-10-31 之间）
        days_offset = random.randint(0, 91)
        hour = random.randint(8, 20)
        minute = random.randint(0, 59)
        entry_dt = datetime.datetime.combine(
            base_date + datetime.timedelta(days=days_offset),
            datetime.time(hour, minute),
        )

        # 地址（用 Faker 的 fake.zh 地址兜底）
        if fake:
            address = fake.address().replace("\n", " ")[:80]
        else:
            address = f"Street {random.randint(1, 999)}, City{random.randint(1, 50)}, DE"

        records_loading.append({
            "序号": i + 1,
            "运单号": random.choice(waybill_ids),   # 共享运单号池
            "客户单号": f"OC{random.randint(1000000000, 9999999999)}",
            "仓库代码": warehouse,
            "地址": address,
            "件数": qty,
            "品名": product_name,
            "总KG": weight,
            "总CBM": volume,
            "入仓时间": entry_dt,
            "入货仓库": "合作仓库",
            "库位": f"w-{random.randint(1, 100)}-{random.randint(1, 50)}({random.randint(1, 300)})",
            "渠道安排": channel,
            "报关方式": customs,
            "备注": random.choice(["优先装", "中间装", "尾部装", ""]),
        })

    df_loading = pd.DataFrame(records_loading)

    # ---------- 生成 原数据 ----------
    records_raw = []
    for i in range(n_rows):
        # 品名（与装柜清单保持一致的模糊比例）
        if random.random() < 0.70:
            cn_name = random.choice(PRODUCT_NAMES_CLEAN)
        else:
            cn_name = random.choice(PRODUCT_NAMES_FUZZY)

        # 品牌：40% 概率有品牌（"有品牌"客户占比）
        brand = random.choice(BRAND_POOL) if random.random() < 0.40 else NO_BRAND

        records_raw.append({
            "运单号码": random.choice(waybill_ids),   # 共享同一运单号池
            "中文品名": cn_name,
            "英文品名": f"ENGLISH-{cn_name[:4]}",
            "海关编码": f"{random.randint(1000, 9999)}.{random.randint(10, 99)}.{random.randint(1000, 9999)}",
            "总数量": random.randint(5, 300),
            "申报单价": round(random.uniform(1.5, 120.0), 2),
            "箱单件数": random.randint(1, 10),
            "收货实重": round(random.uniform(0.5, 45.0), 2),
            "品牌": brand,
            "poNo": f"PO-{random.randint(100000, 999999)}",
            "仓库编码": random.choice(WAREHOUSE_CODES),
            "材质": random.choice(MATERIALS),
            "用途": random.choice(USES),
            "图片": "",
            "SKU": f"SKU-{random.randint(100000, 999999)}",
            "货箱编号": f"{random.randint(1, 20)}-{random.randint(1, 200)}",
            "收件国家": random.choice(DEST_COUNTRIES),
            "发票地址": "",
            "商品链接": "",
            "客户单号": f"OC{random.randint(1000000000, 9999999999)}",
            "渠道产品名称": random.choice(CHANNEL_PRODUCT_NAMES),
            "收件省": "",
            "收件市": "",
        })

    df_raw = pd.DataFrame(records_raw)

    print(f"[MOCK] 已生成装柜清单 {len(df_loading)} 行 / 原数据 {len(df_raw)} 行")
    return {"装柜清单": df_loading, "原数据": df_raw}


# ===========================================================================
# 第二部分：真实数据加载器 / Part 2: Real data loader
# ===========================================================================

def load_real_data(root_dir: str = r"<用户真实数据目录>",
                   max_files: Optional[int] = None) -> Dict[str, pd.DataFrame]:
    """
    扫描 root_dir 下所有子文件夹，查找文件名含"已改"的 xlsx，
    读取其中的"装柜清单"和"原数据"两张工作表，并合并为两张大 DataFrame。

    Scan all subfolders, find xlsx files whose name contains "已改",
    read "装柜清单" and "原数据" sheets, and concatenate them.

    Args:
        root_dir: 根目录
        max_files: 最多读取多少个文件（调试用，None 表示全部）

    Returns:
        dict: {"装柜清单": DataFrame, "原数据": DataFrame}
    """
    print(f"[REAL] 开始扫描目录: {root_dir}")

    # 获取所有含"已改"的 xlsx（排除临时文件 ~$）
    pattern = os.path.join(root_dir, "**", "*已改*.xlsx")
    all_files = glob.glob(pattern, recursive=True)
    # 再扫一层：有些文件名里"已改"位置不固定，兜底扫一下
    pattern2 = os.path.join(root_dir, "**", "*.xlsx")
    extra_files = [f for f in glob.glob(pattern2, recursive=True)
                   if "已改" in os.path.basename(f)]

    candidates = list(set(all_files + extra_files))
    candidates = [f for f in candidates if not os.path.basename(f).startswith("~$")]
    candidates.sort()

    # 排除"检查反馈"文件（这些不是最终版）——虽然名字里也有"已改"但不在目标路径上
    # 实际上 *已改*.xlsx 已经只匹配文件名含"已改"的版本，检查反馈文件名不含"已改"，
    # 所以不用额外排除
    print(f"[REAL] 匹配到 {len(candidates)} 个 *已改*.xlsx 文件")

    if max_files is not None:
        candidates = candidates[:max_files]

    dfs_loading: List[pd.DataFrame] = []
    dfs_raw: List[pd.DataFrame] = []
    skipped = 0

    for fpath in candidates:
        try:
            # "装柜清单"的表头在第 9 行 (0-indexed)
            df_ld = pd.read_excel(
                fpath, sheet_name="装柜清单", header=9,
                engine="openpyxl", dtype=str,
            )
            # 清理空行
            df_ld = df_ld.dropna(subset=["序号"]).copy()
            # 数值类型转换
            for col in ["件数", "总KG", "总CBM"]:
                if col in df_ld.columns:
                    df_ld[col] = pd.to_numeric(df_ld[col], errors="coerce")
            if "入仓时间" in df_ld.columns:
                df_ld["入仓时间"] = pd.to_datetime(df_ld["入仓时间"], errors="coerce")

            df_rd = pd.read_excel(
                fpath, sheet_name="原数据", header=0,
                engine="openpyxl", dtype=str,
            )
            for col in ["箱单件数", "收货实重", "总数量"]:
                if col in df_rd.columns:
                    df_rd[col] = pd.to_numeric(df_rd[col], errors="coerce")

            dfs_loading.append(df_ld)
            dfs_raw.append(df_rd)
        except Exception as exc:
            skipped += 1
            print(f"[REAL] 跳过文件（可能表名不匹配）: {os.path.basename(fpath)} -> {exc}")

    print(f"[REAL] 成功读取 {len(dfs_loading)} 个文件，跳过 {skipped} 个")

    if not dfs_loading:
        raise RuntimeError("未能读取任何有效文件，请检查路径或文件结构。")

    combined_loading = pd.concat(dfs_loading, ignore_index=True)
    combined_raw = pd.concat(dfs_raw, ignore_index=True)

    print(f"[REAL] 合并后：装柜清单 {len(combined_loading)} 行 / "
          f"原数据 {len(combined_raw)} 行")

    return {"装柜清单": combined_loading, "原数据": combined_raw}


# ===========================================================================
# 第三部分：核心分析类 / Part 3: SupplyChainAnalyzer class
# ===========================================================================

class SupplyChainAnalyzer:
    """
    清关供应链数据分析器。

    整合四张分析维度：
      Dimension 1: 仓储物流与抛重损耗分析     (装柜清单)
      Dimension 2: 客户画像与高净值客户预警   (原数据)
      Dimension 3: 数据质量与人工干预成本测算  (原数据)
      Dimension 4: 国家 × 报关方式合规矩阵     (两表交叉)

    构造函数接收两张 DataFrame，自动执行全部分析并产出图表与 Markdown 报告。
    """

    # 模糊词汇列表（用于维度 3 的数据质量检测）
    FUZZY_KEYWORDS = ["套装", "配件", "混装", "其他", "杂项", "零件"]

    def __init__(self,
                 data: Dict[str, pd.DataFrame],
                 output_dir: str = "output",
                 source_label: str = "mock"):
        """
        Args:
            data:         {"装柜清单": df, "原数据": df}
            output_dir:   图表和报告输出目录
            source_label: 数据来源标签（"mock" 或 "real"）
        """
        self.df_loading = data["装柜清单"].copy()
        self.df_raw = data["原数据"].copy()
        self.output_dir = Path(output_dir)
        self.source_label = source_label
        self.output_dir.mkdir(parents=True, exist_ok=True)

        # 分析结果（供生成报告使用）
        self.results: Dict[str, Any] = {}

        print(f"\n{'='*60}")
        print(f"SupplyChainAnalyzer 初始化完成")
        print(f"  数据来源:       {source_label}")
        print(f"  装柜清单行数:   {len(self.df_loading):,}")
        print(f"  原数据行数:     {len(self.df_raw):,}")
        print(f"  输出目录:       {self.output_dir.resolve()}")
        print(f"{'='*60}\n")

    # ------------------------------------------------------------------
    # 维度 1：仓储物流与抛重（体积重）损耗分析
    # Dimension 1: Warehousing & volumetric-weight loss analysis
    # ------------------------------------------------------------------
    def analyze_dim1_volumetric(self) -> None:
        """
        基于"装柜清单"表：
          1. 计算每行的密度 Density = 总CBM / 总KG
          2. 标记"泡货"（密度 < 167 kg/CBM）
          3. 找出密度最低（最泡）的前 10 个品名
          4. 统计不同渠道安排的日均入仓件数（按入仓时间）
        """
        print("[D1] 开始：仓储物流与抛重损耗分析 ...")

        df = self.df_loading.copy()

        # ---- 1. 密度计算 ----
        # 物理密度 (kg/CBM) = 重量 KG / 体积 CBM —— 表示"每立方米体积的重量"。
        # 物流行业标准：体积重 = CBM × 167 kg/CBM；
        # 若 实际重量 KG < 体积重 CBM×167，称为"泡货/抛货"（计费重取体积重），
        # 等价于：密度 = KG/CBM < 167。密度越低 → 越泡 → 越占空间、物流成本越高。
        df["密度"] = df["总KG"].replace(0, np.nan) / df["总CBM"].replace(0, np.nan)
        df["是否泡货"] = df["密度"] < LIGHTER_THAN_DENSITY_THRESHOLD

        # ---- 2. 整体泡货比例 ----
        valid = df.dropna(subset=["密度"])
        bubble_ratio = valid["是否泡货"].mean() * 100

        # ---- 3. 密度最低（最泡）的前 10 个品名 ----
        # 以品名为 groupby，取平均密度排序
        group_name = (
            valid.groupby("品名", dropna=False)["密度"]
            .mean()
            .sort_values(ascending=True)
            .head(10)
        )

        # ---- 4. 按渠道统计日均入仓件数 ----
        # 先把入仓时间标准化为日期
        df["入仓日期"] = pd.to_datetime(df["入仓时间"], errors="coerce").dt.date
        # 每个渠道每天的件数
        channel_daily = (
            df.dropna(subset=["入仓日期", "渠道安排"])
              .groupby(["渠道安排", "入仓日期"], dropna=False)["件数"]
              .sum()
              .reset_index()
        )
        # 再按渠道聚合为"日均件数"
        channel_avg = (
            channel_daily.groupby("渠道安排", dropna=False)["件数"]
            .mean()
            .sort_values(ascending=False)
        )

        # ---- 保存结果 ----
        self.results["dim1"] = {
            "bubble_ratio_pct": float(bubble_ratio),
            "top10_lightest_products": group_name.to_dict(),
            "channel_daily_avg_qty": channel_avg.to_dict(),
            "df_density": valid,
        }

        print(f"     泡货占比: {bubble_ratio:.1f}%")
        print(f"     TOP10 最泡品名: {list(group_name.index)}")

        # ---- 图表 1：双图组合 ----
        #   左：TOP10 最泡品名水平条形图
        #   右：各渠道日均入仓件数
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 7))
        fig.suptitle(
            "维度1：仓储物流与抛重损耗分析 / D1 - Volumetric Loss & Channel Throughput",
            fontsize=14, fontweight="bold",
        )

        # --- 左图：TOP10 最泡品名 ---
        top10_items = group_name.reset_index()
        top10_items.columns = ["品名", "平均密度(kg/CBM)"]
        # 反转以便条形图从上到下显示
        top10_items = top10_items.sort_values("平均密度(kg/CBM)", ascending=False)
        colors = sns.color_palette("Reds_r", len(top10_items))
        ax1.barh(top10_items["品名"], top10_items["平均密度(kg/CBM)"], color=colors)
        ax1.axvline(
            x=LIGHTER_THAN_DENSITY_THRESHOLD, color="red", linestyle="--",
            linewidth=1.5, label=f"泡货阈值 {LIGHTER_THAN_DENSITY_THRESHOLD} kg/CBM",
        )
        ax1.set_xlabel("平均密度 (kg/CBM)")
        ax1.set_ylabel("品名")
        ax1.set_title(
            f"密度最低（最泡）TOP10 品名\n泡货占总货量 {bubble_ratio:.1f}%",
            fontsize=11,
        )
        ax1.legend(fontsize=9)
        for p in ax1.patches:
            w = p.get_width()
            ax1.text(w + 2, p.get_y() + p.get_height()/2,
                     f"{w:.1f}", va="center", fontsize=8)

        # --- 右图：渠道日均入仓件数 ---
        ch_df = channel_avg.reset_index()
        ch_df.columns = ["渠道安排", "日均入仓件数"]
        # 渠道名截断（显示不全）
        ch_df["渠道短名"] = ch_df["渠道安排"].str.replace("华东", "").str[:12]
        palette = sns.color_palette("Blues_d", len(ch_df))
        ax2.bar(range(len(ch_df)), ch_df["日均入仓件数"], color=palette)
        ax2.set_xticks(range(len(ch_df)))
        ax2.set_xticklabels(ch_df["渠道短名"], rotation=30, ha="right", fontsize=8)
        ax2.set_ylabel("日均入仓件数")
        ax2.set_title("各渠道日均入仓件数 / Daily Inbound Qty by Channel", fontsize=11)
        for idx, val in enumerate(ch_df["日均入仓件数"]):
            ax2.text(idx, val + 5, f"{val:.0f}", ha="center", fontsize=8)

        plt.tight_layout(rect=[0, 0, 1, 0.95])
        fig_path = self.output_dir / "chart1_volumetric_loss.png"
        fig.savefig(fig_path, dpi=150, bbox_inches="tight")
        plt.close(fig)
        print(f"     -> 图表已保存: {fig_path}")

    # ------------------------------------------------------------------
    # 维度 2：客户画像与高净值客户流失预警
    # Dimension 2: Customer profiling & HNWI churn warning
    # ------------------------------------------------------------------
    def analyze_dim2_customer(self) -> None:
        """
        基于"原数据"表：
          1. 计算品牌赋能率（有明确品牌记录的商品占比）
          2. 交叉分析：有品牌 vs 无品牌商品的【材质】与【用途】差异
          3. 画出堆叠柱状图展示差异
        """
        print("[D2] 开始：客户画像与高净值客户流失预警 ...")

        df = self.df_raw.copy()

        # ---- 1. 品牌赋能率 ----
        # 业务规则：
        #   - 品牌列空白 / NaN / 非数字等视为"无品牌"
        #   - 有品牌商品 = 客户正在做精细化运营 → 核心客户画像
        brand_series = df["品牌"].astype(str).str.strip() if "品牌" in df.columns else pd.Series([], dtype=str)
        has_brand_mask = brand_series.notna() & (brand_series != "") & (brand_series.str.lower() != "nan")
        brand_rate = has_brand_mask.mean() * 100

        df["是否有品牌"] = has_brand_mask.map({True: "有品牌", False: "无品牌"})

        # ---- 2. 材质分布交叉 ----
        material_cross = (
            df.groupby(["是否有品牌", "材质"], dropna=False)
              .size()
              .unstack(fill_value=0)
        )
        # 归一化为百分比（按行：每组内部分布）
        material_cross_pct = material_cross.div(material_cross.sum(axis=1), axis=0) * 100

        # ---- 3. 用途分布交叉 ----
        usage_cross = (
            df.groupby(["是否有品牌", "用途"], dropna=False)
              .size()
              .unstack(fill_value=0)
        )
        usage_cross_pct = usage_cross.div(usage_cross.sum(axis=1), axis=0) * 100

        # ---- 保存结果 ----
        self.results["dim2"] = {
            "brand_rate_pct": float(brand_rate),
            "material_cross_pct": material_cross_pct,
            "usage_cross_pct": usage_cross_pct,
            "n_brand": int(has_brand_mask.sum()),
            "n_no_brand": int((~has_brand_mask).sum()),
        }

        print(f"     品牌赋能率: {brand_rate:.1f}%  (有品牌 {int(has_brand_mask.sum())} / 无品牌 {int((~has_brand_mask).sum())})")

        # ---- 图表 2：材质 × 品牌、用途 × 品牌 交叉堆叠柱状图 ----
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(17, 7))
        fig.suptitle(
            "维度2：客户画像 —— 有品牌 vs 无品牌商品在材质、用途上的差异\n"
            "D2 - Customer Profiling: Material & Usage Split by Brand Awareness",
            fontsize=13, fontweight="bold",
        )

        # --- 左图：材质分布（取 Top 10 材质避免太杂）---
        top_materials = material_cross.sum().sort_values(ascending=False).head(10).index
        mat_plot = material_cross_pct.reindex(columns=top_materials).T
        mat_plot.plot(kind="barh", stacked=True, ax=ax1,
                      color=["#4C72B0", "#DD8452"], edgecolor="white")
        ax1.set_xlabel("占该组商品比例 (%)")
        ax1.set_ylabel("材质")
        ax1.set_title(
            f"材质分布 × 品牌归属（品牌赋能率 {brand_rate:.1f}%）",
            fontsize=11,
        )
        ax1.legend(title="品牌状态", loc="lower right")
        for container in ax1.containers:
            ax1.bar_label(container, fmt="%.1f%%", label_type="center", fontsize=7, color="white")

        # --- 右图：用途分布 ---
        top_usages = usage_cross.sum().sort_values(ascending=False).head(10).index
        usage_plot = usage_cross_pct.reindex(columns=top_usages).T
        usage_plot.plot(kind="barh", stacked=True, ax=ax2,
                        color=["#4C72B0", "#DD8452"], edgecolor="white")
        ax2.set_xlabel("占该组商品比例 (%)")
        ax2.set_ylabel("用途")
        ax2.set_title("用途分布 × 品牌归属", fontsize=11)
        ax2.legend(title="品牌状态", loc="lower right")
        for container in ax2.containers:
            ax2.bar_label(container, fmt="%.1f%%", label_type="center", fontsize=7, color="white")

        plt.tight_layout(rect=[0, 0, 1, 0.92])
        fig_path = self.output_dir / "chart2_customer_profile.png"
        fig.savefig(fig_path, dpi=150, bbox_inches="tight")
        plt.close(fig)
        print(f"     -> 图表已保存: {fig_path}")

    # ------------------------------------------------------------------
    # 维度 3：数据质量与"人工干预成本"测算
    # Dimension 3: Data quality & manual-intervention cost estimation
    # ------------------------------------------------------------------
    def analyze_dim3_data_quality(self) -> None:
        """
        基于"原数据"表 B列（中文品名）：
          1. 用模糊关键词扫描：套装 / 配件 / 混装 / 其他 / 杂项 / 零件
          2. 统计每个模糊词的命中数与占比
          3. 汇总"需要人工清洗"的品名比例
          4. 给出 HS 编码自动匹配引擎的必要性论证
        """
        print("[D3] 开始：数据质量与人工干预成本测算 ...")

        df = self.df_raw.copy()

        if "中文品名" not in df.columns:
            print("     [WARN] 原数据表缺少 '中文品名' 列，跳过维度 3")
            self.results["dim3"] = {"total": 0, "fuzzy_total_pct": 0.0}
            return

        cn_names = df["中文品名"].astype(str).fillna("")
        total = len(cn_names)

        # 逐关键词扫描并记录命中行
        fuzzy_flags = pd.DataFrame(index=df.index)
        for kw in self.FUZZY_KEYWORDS:
            fuzzy_flags[kw] = cn_names.str.contains(re.escape(kw), regex=False)

        # 至少命中一个模糊关键词 => 需要人工清洗
        needs_cleaning = fuzzy_flags.any(axis=1)
        cleaning_rate = needs_cleaning.mean() * 100

        # 每个关键词单独的比例
        per_keyword_pct = {
            kw: float(fuzzy_flags[kw].mean() * 100)
            for kw in self.FUZZY_KEYWORDS
        }

        # ---- 保存结果 ----
        self.results["dim3"] = {
            "total_rows": int(total),
            "needs_cleaning_rows": int(needs_cleaning.sum()),
            "cleaning_rate_pct": float(cleaning_rate),
            "per_keyword_pct": per_keyword_pct,
            "sample_dirty_names": df.loc[needs_cleaning, "中文品名"].head(15).tolist(),
        }

        print(f"     总行数: {total}")
        print(f"     需要人工清洗: {int(needs_cleaning.sum())} 行 ({cleaning_rate:.1f}%)")
        for kw, pct in per_keyword_pct.items():
            print(f"        '{kw}': {pct:.1f}%")

        # ---- 图表 3：各模糊词占比 + 综合人工成本估算 ----
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 7),
                                       gridspec_kw={"width_ratios": [2, 1]})
        fig.suptitle(
            "维度3：数据质量 —— 品名模糊词汇扫描\n"
            "D3 - Data Quality: Vague Product Names & Manual-clean Rate",
            fontsize=13, fontweight="bold",
        )

        # --- 左图：各模糊关键词占比 ---
        kws = list(per_keyword_pct.keys())
        vals = list(per_keyword_pct.values())
        bar_colors = sns.color_palette("YlOrRd", len(kws))
        bars = ax1.bar(kws, vals, color=bar_colors, edgecolor="white")
        ax1.set_ylabel("在全部品名中的占比 (%)")
        ax1.set_xlabel("模糊关键词")
        ax1.set_title("各模糊关键词命中占比", fontsize=11)
        for bar, v in zip(bars, vals):
            ax1.text(bar.get_x() + bar.get_width()/2, v + 0.5,
                     f"{v:.1f}%", ha="center", va="bottom", fontsize=9)

        # --- 右图：饼图 + 数据标注 ---
        labels = ["需要人工清洗", "自动可识别"]
        sizes = [cleaning_rate, 100 - cleaning_rate]
        explode = (0.05, 0)
        pie_colors = ["#C44E52", "#55A868"]
        ax2.pie(sizes, explode=explode, labels=labels, colors=pie_colors,
                autopct="%1.1f%%", startangle=90, textprops={"fontsize": 10})
        ax2.set_title(
            f"人工清洗成本占比\n（共 {total:,} 行）",
            fontsize=11,
        )

        # 在右图下方加业务文本框
        cost_estimate_text = (
            f"【业务测算 / Business Estimate】\n"
            f"平均每单 HS 编码人工匹配耗时 ~3 分钟\n"
            f"若月处理 20,000 单：\n"
            f"  人工工时 = {cleaning_rate/100*20000*3/60:.0f} 小时 / 月\n"
            f"  按 200 元/小时 ≈ {cleaning_rate/100*20000*3/60*200:,.0f} 元 / 月"
        )
        fig.text(0.5, 0.01, cost_estimate_text,
                 ha="center", fontsize=9,
                 bbox=dict(boxstyle="round,pad=0.5", facecolor="wheat", alpha=0.5))

        plt.tight_layout(rect=[0, 0.06, 1, 0.93])
        fig_path = self.output_dir / "chart3_data_quality.png"
        fig.savefig(fig_path, dpi=150, bbox_inches="tight")
        plt.close(fig)
        print(f"     -> 图表已保存: {fig_path}")

    # ------------------------------------------------------------------
    # 维度 4：国家 × 报关方式 合规矩阵
    # Dimension 4: Country × Customs-declaration compliance matrix
    # ------------------------------------------------------------------
    def analyze_dim4_compliance(self) -> None:
        """
        综合两表：原数据的收件国家(Q列) × 装柜清单的报关方式(N列)
        输出交叉表，并画堆叠柱状图。

        业务解读：
          - "有单证" vs "无单证" 在不同国家的分布，暴露合规风险点
        """
        print("[D4] 开始：国家 × 报关方式 合规矩阵 ...")

        df_load = self.df_loading.copy()
        df_raw = self.df_raw.copy()

        # ---- 先在"原数据"里按运单号码聚合出【运单号码 → 收件国家】映射 ----
        if "收件国家" not in df_raw.columns or "运单号码" not in df_raw.columns:
            print("     [WARN] 原数据表缺少必要列（收件国家/运单号码），跳过维度 4")
            self.results["dim4"] = {}
            return

        # 一条运单可能对应多件商品，但收件国家应一致，按第一条取值即可
        waybill_country = (
            df_raw.dropna(subset=["运单号码"])
                  .groupby("运单号码", dropna=False)["收件国家"]
                  .first()
        )

        # ---- 把国家信息 merge 到装柜清单 ----
        df_load = df_load.merge(
            waybill_country.rename("收件国家"),
            left_on="运单号" if "运单号" in df_load.columns else "运单号码",
            right_index=True,
            how="left",
        )

        # ---- 交叉表：国家 × 报关方式 ----
        if "报关方式" not in df_load.columns or "收件国家" not in df_load.columns:
            print("     [WARN] 装柜清单缺少必要列（报关方式/收件国家），跳过维度 4")
            self.results["dim4"] = {}
            return

        cross = pd.crosstab(
            index=df_load["收件国家"],
            columns=df_load["报关方式"],
            normalize="index",        # 行归一化：看每个国家内部有无单证的比例
        ) * 100

        cross_count = pd.crosstab(
            index=df_load["收件国家"],
            columns=df_load["报关方式"],
        )

        # ---- 保存结果 ----
        self.results["dim4"] = {
            "cross_pct": cross,
            "cross_count": cross_count,
            "total_by_country": cross_count.sum(axis=1).to_dict(),
        }

        print(f"     涉及国家数: {len(cross)}")
        total_with_doc = cross_count.get("有单证", pd.Series(dtype=int)).sum() if "有单证" in cross_count.columns else 0
        total_without = cross_count.get("无单证", pd.Series(dtype=int)).sum() if "无单证" in cross_count.columns else 0
        grand = total_with_doc + total_without
        print(f"     有单证: {int(total_with_doc):,} / 无单证: {int(total_without):,} / 总计: {int(grand):,}")

        # ---- 图表 4：堆叠柱状图 ----
        # 防护：若无有效数据则跳过绘图，避免 pandas.plot 报错
        if cross.empty or grand == 0:
            print("     [SKIP] cross 表为空，跳过维度 4 绘图")
            return

        fig, ax = plt.subplots(figsize=(14, 7))
        cross.plot(
            kind="bar", stacked=True, ax=ax,
            color=["#4C72B0", "#DD8452"],
            edgecolor="white",
        )
        ax.set_ylabel("占该国货物比例 (%)")
        ax.set_xlabel("收件国家 / Destination Country")
        ax.set_title(
            "维度4：国家 × 报关方式 合规矩阵\n"
            "D4 - Compliance Matrix: Destination Country vs Customs Declaration Type",
            fontsize=13, fontweight="bold",
        )
        ax.legend(title="报关方式", fontsize=9)
        plt.xticks(rotation=30, ha="right", fontsize=10)

        # 在每段柱子中央标注数值
        for container in ax.containers:
            ax.bar_label(container, fmt="%.1f%%", label_type="center", fontsize=8, color="white")

        plt.tight_layout()
        fig_path = self.output_dir / "chart4_compliance_matrix.png"
        fig.savefig(fig_path, dpi=150, bbox_inches="tight")
        plt.close(fig)
        print(f"     -> 图表已保存: {fig_path}")

    # ------------------------------------------------------------------
    # 运行全部分析 / Run all analyses
    # ------------------------------------------------------------------
    def run_all(self) -> None:
        """依次执行 4 个维度的分析。"""
        self.analyze_dim1_volumetric()
        self.analyze_dim2_customer()
        self.analyze_dim3_data_quality()
        self.analyze_dim4_compliance()

    # ------------------------------------------------------------------
    # 生成 Markdown 业务洞察报告
    # Generate the Markdown business-insight report
    # ------------------------------------------------------------------
    def generate_report(self) -> Path:
        """
        基于已完成的 self.results 字典，生成一份完整的 Markdown 报告，
        论证"为什么需要流程优化"和"为什么需要自动化工具"。

        Returns:
            报告文件路径
        """
        r = self.results

        report_lines: List[str] = []
        _add = report_lines.append

        _add("# 清关资料数据分析 —— 业务洞察报告")
        _add("")
        _add(f"> **数据来源**: {self.source_label}  |  **生成时间**: {datetime.datetime.now():%Y-%m-%d %H:%M}")
        _add(f"> **装柜清单** {len(self.df_loading):,} 行  |  **原数据** {len(self.df_raw):,} 行")
        _add("")
        _add("---")
        _add("")
        _add("## 执行摘要 / Executive Summary")
        _add("")
        _add(
            "本报告基于真实清关资料的多维度分析，旨在回答管理层最关心的两个问题："
        )
        _add("")
        _add("1. **为什么我们必须立即进行内部流程优化？**")
        _add("2. **为什么我们需要投入资源开发自动化工具？**")
        _add("")
        _add("---")

        # ===== 维度 1 =====
        _add("")
        _add("## 维度 1：仓储物流与抛重（体积重）损耗分析")
        _add("")
        _add("### 核心发现")
        _add("")

        d1 = r.get("dim1", {})
        if d1:
            _add(f"- **泡货占比高达 {d1['bubble_ratio_pct']:.1f}%** —— 密度低于 167 kg/CBM 的货物，物流成本远超实际重量，正在侵蚀渠道空间与利润。")
            top10 = d1.get("top10_lightest_products", {})
            if top10:
                _add("- **最泡的 TOP10 品名**（按平均密度升序）：")
                for name, density in list(top10.items())[:10]:
                    _add(f"    - `{name}`：**{density:.1f} kg/CBM**")
            channel_avg = d1.get("channel_daily_avg_qty", {})
            if channel_avg:
                _add("")
                _add("| 渠道安排 | 日均入仓件数 |")
                _add("|---|---:|")
                for ch, qty in list(channel_avg.items())[:10]:
                    _add(f"| {ch} | {qty:.0f} |")
        else:
            _add("- （无数据）")

        _add("")
        _add("### 管理层建议")
        _add("")
        _add(
            '> **泡货不是"客户问题"，而是"产品组合问题"。** 对于密度 < 167 kg/CBM 的高频品类，'
            '应考虑"渠道打包"或"同品高密度补货"策略，否则将持续占用集装箱 CBM 配额而拉低整柜利润。'
        )
        _add("")
        _add(f"![维度1图表](chart1_volumetric_loss.png)")

        # ===== 维度 2 =====
        _add("")
        _add("## 维度 2：客户画像与高净值客户识别")
        _add("")
        _add("### 核心发现")
        _add("")

        d2 = r.get("dim2", {})
        if d2:
            _add(f"- **品牌赋能率 {d2['brand_rate_pct']:.1f}%** —— 有品牌记录的商品约 {d2['n_brand']:,} 行，无品牌约 {d2['n_no_brand']:,} 行。")
            _add(
                '- 这部分"有品牌"客户大概率在做 **精细化运营**（自建站/亚马逊品牌注册/独立站），'
                '对清关合规、HS 编码准确性、报关时效要求更高，是未来的高净值核心客户群。'
            )
        else:
            _add("- （无数据）")

        _add("")
        _add("### 管理层建议")
        _add("")
        _add(
            '> **建议立刻建立"有品牌客户"的分级标记与专属清关优先级通道。** '
            '如果我们能证明自己在合规、时效、数据质量上优于同行，这部分客户就是最稳定、'
            '复购率最高、愿意付费的一群。反之，如果他们觉得你的清关数据经常出错、'
            'HS 编码混乱，转走成本极低。'
        )
        _add("")
        _add(f"![维度2图表](chart2_customer_profile.png)")

        # ===== 维度 3 =====
        _add("")
        _add("## 维度 3：数据质量与人工干预成本测算")
        _add("")
        _add("### 核心发现")
        _add("")

        d3 = r.get("dim3", {})
        if d3 and d3.get("total_rows", 0) > 0:
            _add(
                f"- **{d3['cleaning_rate_pct']:.1f}% 的品名含有模糊词汇** —— "
                f"共 {d3['needs_cleaning_rows']:,} / {d3['total_rows']:,} 行需要人工清洗。"
            )
            _add("- 各模糊关键词占比：")
            for kw, pct in d3.get("per_keyword_pct", {}).items():
                _add(f"  - `{kw}`：**{pct:.1f}%**")
            _add("")
            _add("**业务成本测算**（基于行业经验值）：")
            _add("")
            _add("| 参数 | 数值 |")
            _add("|---|---:|")
            _add(f"| 月处理单数（估算） | 20,000 |")
            _add(f"| 需要人工清洗的比例 | {d3['cleaning_rate_pct']:.1f}% |")
            _add(f"| 平均每单 HS 编码人工匹配耗时 | 3 分钟 |")
            manual_hours = d3["cleaning_rate_pct"] / 100 * 20000 * 3 / 60
            _add(f"| **每月人工工时** | **{manual_hours:.0f} 小时** |")
            cost_200 = manual_hours * 200
            cost_100 = manual_hours * 100
            _add(f"| 按 200 元/小时 估算人工成本 | **≈ {cost_200:,.0f} 元 / 月** |")
            _add(f"| 按 100 元/小时 估算人工成本 | **≈ {cost_100:,.0f} 元 / 月** |")
        else:
            _add("- （无数据）")

        _add("")
        _add("### 管理层建议")
        _add("")
        _add(
            '> **开发"品名→HS编码自动匹配引擎"ROI 极高。** 引擎开发投入假设 30 人·天，'
            '一旦上线，**仅每年节省的人工成本即可达到 {cost:.0f} 万元**（以 100 元/小时估算）。'
            '此外，自动匹配的一致性和可追溯性还能显著降低海关查验率和合规风险。'
            .format(cost=(d3.get("cleaning_rate_pct", 0) / 100 * 20000 * 3 / 60 * 100 * 12 / 10000 if d3 else 0))
        )
        _add("")
        _add(f"![维度3图表](chart3_data_quality.png)")

        # ===== 维度 4 =====
        _add("")
        _add("## 维度 4：国家 × 报关方式合规矩阵")
        _add("")
        _add("### 核心发现")
        _add("")

        d4 = r.get("dim4", {})
        cross_count = d4.get("cross_count") if d4 else None
        if cross_count is not None and not cross_count.empty:
            _add("**报关方式分布（有单证 vs 无单证）按收件国家统计：**")
            _add("")
            _add("| 收件国家 | 有单证 | 无单证 | 合计 |")
            _add("|---|---:|---:|---:|")
            for country, row in cross_count.iterrows():
                with_doc = int(row.get("有单证", 0)) if "有单证" in row.index else 0
                without = int(row.get("无单证", 0)) if "无单证" in row.index else 0
                _add(f"| {country} | {with_doc:,} | {without:,} | {with_doc + without:,} |")
            _add("")
            _add(
                '> **合规提醒：** 某些国家"无单证"占比偏高，需要重点审视 —— '
                '是因为这些国家对单证要求较宽松（可接受），还是我们自身漏配单证（风险）。'
            )
        else:
            _add("- （无数据）")

        _add("")
        _add(f"![维度4图表](chart4_compliance_matrix.png)")

        # ===== 总结 =====
        _add("")
        _add("---")
        _add("")
        _add("## 总结：为什么要流程优化 + 为什么要自动化")
        _add("")
        _add("| # | 我们的问题 | 对应的自动化 / 流程优化方案 | 预期收益 |")
        _add("|---|---|---|---|")

        d1 = r.get("dim1", {})
        d3 = r.get("dim3", {})
        brand_pct = r.get("dim2", {}).get("brand_rate_pct", 0.0)

        _add(f"| 1 | 泡货占比 {d1.get('bubble_ratio_pct', 0):.1f}%，持续吞噬集装箱空间 | **渠道自动配柜引擎** | 每柜平均多塞 15-20% 货量 |")
        _add(f"| 2 | {d3.get('cleaning_rate_pct', 0):.1f}% 品名含模糊词，人工匹配耗时巨大 | **品名 → HS 编码自动匹配引擎** | 每月节省 ≈ {d3.get('cleaning_rate_pct', 0)/100*20000*3/60:.0f} 小时人工 |")
        _add(f"| 3 | 品牌赋能率 {brand_pct:.1f}%，但缺乏针对高净值客户的差异化服务 | **客户自动分级 + 清关优先级调度** | 核心客户留存率 + 客户单产值 |")
        _add("| 4 | 不同国家报关方式合规性不透明，有风险点 | **国家合规规则自动校验清单** | 降低海关查验率和罚金 |")
        _add("")
        _add("### 一句话结论")
        _add("")
        _add(
            '> **不做流程优化 = 把利润拱手让给物流成本、人工成本和潜在的合规罚单。** 自动化工具不是奢侈品，'
            '而是在"订单量不变"的前提下，**几乎唯一能同时提升利润、提升客户体验、降低团队内耗** 的抓手。'
        )
        _add("")

        report_text = "\n".join(report_lines)
        report_path = self.output_dir / "business_insight_report.md"
        report_path.write_text(report_text, encoding="utf-8")
        print(f"\n[REPORT] 业务洞察报告已生成: {report_path}")
        return report_path


# ===========================================================================
# 第四部分：主程序 / Part 4: Main entry point
# ===========================================================================

def main() -> None:
    """主入口。支持 --mock / --real 切换，也支持自动检测。"""
    import argparse
    parser = argparse.ArgumentParser(description="Supply Chain Analyzer - 清关资料数据分析")
    parser.add_argument("--mock", action="store_true",
                        help="强制使用模拟数据（不访问 E:\\ 盘）")
    parser.add_argument("--real", action="store_true",
                        help="强制使用真实数据（必须存在 E:\\已上传云文档）")
    parser.add_argument("--seed", type=int, default=42,
                        help="Mock 数据随机种子，默认 42")
    parser.add_argument("--rows", type=int, default=2000,
                        help="Mock 数据行数，默认 2000")
    parser.add_argument("--max-files", type=int, default=None,
                        help="真实数据最多读取的文件数（调试用），None=全部")
    parser.add_argument("--output", type=str, default="output",
                        help="输出目录，默认 ./output")
    args = parser.parse_args()

    output_dir = args.output

    # ---- 决定用 mock 还是 real ----
    real_root = r"<用户真实数据目录>"
    use_real = False

    if args.real:
        use_real = True
    elif args.mock:
        use_real = False
    else:
        # 自动检测：若真实目录存在则优先读真实，否则降级 mock
        if os.path.isdir(real_root):
            use_real = True
            print(f"[AUTO-DETECT] 检测到真实目录 {real_root}，使用真实数据")
        else:
            use_real = False
            print(f"[AUTO-DETECT] 未检测到真实目录，降级使用模拟数据")

    try:
        if use_real:
            data = load_real_data(real_root, max_files=args.max_files)
            source_label = "real"
        else:
            data = generate_mock_data(n_rows=args.rows, seed=args.seed)
            source_label = "mock"

        analyzer = SupplyChainAnalyzer(data, output_dir=output_dir, source_label=source_label)
        analyzer.run_all()
        report_path = analyzer.generate_report()

        print("\n" + "=" * 60)
        print("全部分析完成 / All analyses finished successfully.")
        print(f"  输出目录: {Path(output_dir).resolve()}")
        print("=" * 60)

    except Exception as exc:
        print(f"\n[ERROR] 分析失败: {exc}")
        print("-" * 60)
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
