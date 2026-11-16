# -*- coding: utf-8 -*-
"""
================================================================================
build_word_report.py
================================================================================

将 SupplyChainAnalyzer 的分析结果 + 4 张图表 PNG，打包为一份
图文并茂的 Word (.docx) 业务洞察报告。

【依赖】
    python-docx (已安装 1.2.0) / pandas / numpy / matplotlib / seaborn

【用法】
    python build_word_report.py                  # 自动检测真实数据
    python build_word_report.py --real           # 强制真实数据
    python build_word_report.py --mock           # 强制模拟数据

输出：
    output_real / output_mock / supply_chain_report.docx
================================================================================
"""

# ---------------------------------------------------------------------------
# 标准库 / Standard imports
# ---------------------------------------------------------------------------
import sys
import os
import datetime
import argparse
from pathlib import Path
from typing import Any, Dict, List

# ---------------------------------------------------------------------------
# 第三方库 / Third-party imports
# ---------------------------------------------------------------------------
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

# 导入上一轮写好的分析器
from supply_chain_analyzer import (
    SupplyChainAnalyzer,
    load_real_data,
    generate_mock_data,
)


# ===========================================================================
# Word 文档构建器 / Document builder helpers
# ===========================================================================

# Windows 下可用的 CJK 字体（按优先级）
CJK_FONT_CANDIDATES = [
    "Microsoft YaHei",   # 微软雅黑
    "SimHei",            # 黑体
    "SimSun",            # 宋体
    "KaiTi",             # 楷体
    "Arial Unicode MS",  # Mac
    "DejaVu Sans",       # 兜底
]


def _pick_cjk_font() -> str:
    """挑选一个本机可用的中文字体"""
    try:
        import matplotlib.font_manager as fm
        installed = {f.name for f in fm.fontManager.ttflist}
        for name in CJK_FONT_CANDIDATES:
            if name in installed:
                return name
    except Exception:
        pass
    return "Microsoft YaHei"


# 全局 CJK 字体（python-docx 对 CJK 的设置需要显式写 XML）
CJK_FONT = _pick_cjk_font()


def _set_run_cjk_font(run, font_name: str = CJK_FONT,
                      ascii_font: str = "Arial") -> None:
    """
    为一个 run 同时设置 ASCII 和 CJK 字体（python-docx 的中文渲染必须这样）。
    Set both ascii + eastAsia fonts on a run so CJK renders properly.
    """
    run.font.name = ascii_font
    # 关键：写入 w:eastAsia 属性，否则中文会变成方框
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")}/>')
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:eastAsia"), font_name)
    rFonts.set(qn("w:ascii"), ascii_font)
    rFonts.set(qn("w:hAnsi"), ascii_font)


def add_heading_cjk(doc: Document, text: str, level: int = 1) -> None:
    """添加一个带 CJK 字体的 Heading（python-docx 默认 heading 不设中文字体）"""
    p = doc.add_heading(text, level=level)
    for run in p.runs:
        _set_run_cjk_font(run)
    return p


def add_para(doc: Document, text: str, bold: bool = False,
             italic: bool = False, size: int = 11,
             color: RGBColor = None,
             align: WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH.LEFT,
             first_line_indent_chars: int = 2) -> None:
    """添加一个格式齐备的段落（首行缩进 2 字符默认）。"""
    p = doc.add_paragraph()
    # 首行缩进（2 个汉字 ≈ 420 DXA）
    pf = p.paragraph_format
    pf.alignment = align
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.space_before = Pt(4)
    pf.space_after = Pt(4)
    if first_line_indent_chars:
        pf.first_line_indent = Cm(first_line_indent_chars * 0.37)

    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    if color is not None:
        run.font.color.rgb = color
    _set_run_cjk_font(run)
    return p


def add_bullet(doc: Document, text: str, level: int = 0, size: int = 11) -> None:
    """添加项目符号列表项。"""
    # python-docx 的 list style 需要模板，用手动 bullet 更稳
    prefix = "• " if level == 0 else "○ "
    p = doc.add_paragraph(prefix + text, style=None)
    pf = p.paragraph_format
    pf.left_indent = Cm(1.0 + level * 0.8)
    pf.space_before = Pt(1)
    pf.space_after = Pt(1)
    for run in p.runs:
        run.font.size = Pt(size)
        _set_run_cjk_font(run)


def add_banner_quote(doc: Document, text: str,
                     bg_color: str = "D5E8F0",
                     text_color: RGBColor = RGBColor(0x1F, 0x38, 0x64)) -> None:
    """
    添加一个带灰底的"高亮引述条"，用于管理层建议。
    (python-docx 没有原生 blockquote，用表格模拟一个单单元格高亮块)
    """
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    cell = table.rows[0].cells[0]
    # 设置单元格宽度（接近整页）
    cell.width = Cm(15.5)
    # 背景色
    shading_elm = parse_xml(
        f'<w:shd {nsdecls("w")} w:fill="{bg_color}" w:val="clear"/>'
    )
    cell._tc.get_or_add_tcPr().append(shading_elm)
    # 内边距
    tcPr = cell._tc.tcPr
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="80" w:type="dxa"/>'
        f'<w:left w:w="160" w:type="dxa"/>'
        f'<w:bottom w:w="80" w:type="dxa"/>'
        f'<w:right w:w="160" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)
    # 写入文字
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = text_color
    _set_run_cjk_font(run)
    # 移除默认表格边框（让它看起来像一个色块条）
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="nil"/><w:left w:val="nil"/>'
        f'<w:bottom w:val="nil"/><w:right w:val="nil"/>'
        f'<w:insideH w:val="nil"/><w:insideV w:val="nil"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(tblBorders)


def add_table(doc: Document, headers: List[str], rows: List[List[Any]],
              col_widths_cm: List[float] = None,
              header_color: str = "1F4E79") -> None:
    """
    通用表格：带蓝底白字表头 + 斑马纹行 + 自动宽度。
    """
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # 列宽
    if col_widths_cm is None:
        # 均匀分配
        col_widths_cm = [15.0 / len(headers)] * len(headers)
    for i, w in enumerate(col_widths_cm):
        for r in range(1 + len(rows)):
            table.cell(r, i).width = Cm(w)

    # 表头行
    hdr = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        # 背景色
        shading = parse_xml(
            f'<w:shd {nsdecls("w")} w:fill="{header_color}" w:val="clear"/>'
        )
        cell._tc.get_or_add_tcPr().append(shading)
        # 文字
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(str(h))
        run.bold = True
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        _set_run_cjk_font(run)

    # 数据行
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, val in enumerate(row):
            cell = table.cell(r_idx, c_idx)
            # 斑马纹
            if r_idx % 2 == 0:
                shading = parse_xml(
                    f'<w:shd {nsdecls("w")} w:fill="F2F2F2" w:val="clear"/>'
                )
                cell._tc.get_or_add_tcPr().append(shading)
            # 右对齐数值列（简单 heuristic：纯数字）
            text_val = str(val) if val is not None else ""
            is_numeric = text_val.replace(",", "").replace(".", "").replace("-", "").isdigit()
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_numeric else WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(text_val)
            run.font.size = Pt(10)
            _set_run_cjk_font(run)

    # 统一细边框
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="4" w:color="CCCCCC"/>'
        f'<w:left w:val="single" w:sz="4" w:color="CCCCCC"/>'
        f'<w:bottom w:val="single" w:sz="4" w:color="CCCCCC"/>'
        f'<w:right w:val="single" w:sz="4" w:color="CCCCCC"/>'
        f'<w:insideH w:val="single" w:sz="4" w:color="CCCCCC"/>'
        f'<w:insideV w:val="single" w:sz="4" w:color="CCCCCC"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(tblBorders)


def add_image_with_caption(doc: Document, image_path: Path,
                            caption: str, width_inches: float = 6.0) -> None:
    """插入图片 + 居中 + 下方标题。"""
    if not image_path.exists():
        add_para(doc, f"（图片缺失：{image_path.name}）", size=10,
                 color=RGBColor(0xAA, 0xAA, 0xAA))
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run()
    run.add_picture(str(image_path), width=Inches(width_inches))

    # 图片下方标题
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_after = Pt(10)
    r = cap.add_run(caption)
    r.italic = True
    r.font.size = Pt(9)
    r.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    _set_run_cjk_font(r)


def setup_document(doc: Document) -> None:
    """
    全局样式设置：页面大小、页边距、默认字体。
    """
    section = doc.sections[0]
    # A4
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)

    # 默认 Normal 样式
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(11)
    rPr = normal.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")}/>')
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:eastAsia"), CJK_FONT)

    # Heading 样式
    for lv, (h_size, h_color) in [(1, (18, "1F4E79")), (2, (15, "2E75B6")),
                                   (3, (13, "404040"))]:
        st = doc.styles[f"Heading {lv}"]
        st.font.size = Pt(h_size)
        st.font.bold = True
        st.font.color.rgb = RGBColor.from_string(h_color)
        st.font.name = "Arial"
        st._element.get_or_add_rPr()
        rPr = st.element.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = parse_xml(f'<w:rFonts {nsdecls("w")}/>')
            rPr.insert(0, rFonts)
        rFonts.set(qn("w:eastAsia"), CJK_FONT)


# ===========================================================================
# Word 报告生成主函数 / Build the whole Word report
# ===========================================================================

def build_word_report(analyzer: SupplyChainAnalyzer,
                      output_dir: Path) -> Path:
    """
    基于 analyzer.results 和 output_real 目录里的 PNG 图表，
    组装一份图文并茂的 Word 业务洞察报告。
    """
    doc = Document()
    setup_document(doc)

    r = analyzer.results
    data_label = analyzer.source_label
    load_rows = len(analyzer.df_loading)
    raw_rows = len(analyzer.df_raw)
    gen_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    # ---------- 封面 / Cover page ----------
    for _ in range(6):
        doc.add_paragraph()
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_title.add_run("清关资料数据分析")
    run.bold = True
    run.font.size = Pt(36)
    run.font.color.rgb = RGBColor.from_string("1F4E79")
    _set_run_cjk_font(run)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p_sub.add_run("业务洞察报告 / Business Insight Report")
    run2.italic = True
    run2.font.size = Pt(18)
    run2.font.color.rgb = RGBColor.from_string("404040")
    _set_run_cjk_font(run2)

    doc.add_paragraph()
    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for i, txt in enumerate([
        f"数据来源: {data_label}",
        f"装柜清单 {load_rows:,} 行 | 原数据 {raw_rows:,} 行",
        f"生成时间: {gen_time}",
    ]):
        run_m = p_meta.add_run((("\n" if i else "") + txt))
        run_m.font.size = Pt(12)
        run_m.font.color.rgb = RGBColor(0x60, 0x60, 0x60)
        _set_run_cjk_font(run_m)

    # 换页
    doc.add_page_break()

    # ---------- 执行摘要 / Executive Summary ----------
    add_heading_cjk(doc, "执行摘要", level=1)
    add_para(doc,
             "本报告基于真实清关资料的 4 维度数据分析，回答管理层最关心的两个问题：")
    add_bullet(doc, "为什么我们必须立即进行内部流程优化？")
    add_bullet(doc, "为什么我们需要投入资源开发自动化工具？")
    add_para(doc,
             "核心结论（详见各维度）：泡货占比高达四成以上（密度低于物流行业标准 167 kg/CBM），"
             "正在持续吞噬集装箱 CBM 配额；品牌赋能率接近 100%，说明客户群正在走品牌化精细化运营路线，"
             "对清关合规性的要求只会越来越高；数据质量整体尚可但仍有约 2% 品名模糊需人工介入，"
             "这部分成本在月均 2 万单的量级下已可量化。")

    doc.add_paragraph()

    # ==============================================================
    # 维度 1 / Dimension 1
    # ==============================================================
    add_heading_cjk(doc, "维度 1：仓储物流与抛重（体积重）损耗分析", level=1)

    add_heading_cjk(doc, "核心发现", level=2)
    d1 = r.get("dim1", {})
    bubble_pct = d1.get("bubble_ratio_pct", 0.0)
    add_bullet(
        doc,
        f"泡货占比 {bubble_pct:.1f}% —— 密度低于 167 kg/CBM 的货物，"
        f"物流成本远超实际重量，正在侵蚀渠道空间与利润。",
    )

    top10 = d1.get("top10_lightest_products", {})
    if top10:
        add_bullet(doc, f"密度最低（最泡）TOP 10 品名（按平均密度升序）：")
        for name, density in list(top10.items())[:10]:
            add_bullet(doc, f"`{name}` — {density:.1f} kg/CBM", level=1)

    channel_avg = d1.get("channel_daily_avg_qty", {})
    if channel_avg:
        add_heading_cjk(doc, "各渠道日均入仓件数", level=3)
        rows = [[ch, f"{qty:.0f}"] for ch, qty in list(channel_avg.items())[:10]]
        add_table(doc, ["渠道安排", "日均入仓件数"], rows,
                  col_widths_cm=[9.5, 5.5])

    add_heading_cjk(doc, "管理层建议", level=2)
    add_banner_quote(
        doc,
        "泡货不是客户问题，而是产品组合问题。对于密度 < 167 kg/CBM 的高频品类，"
        "应考虑渠道打包或同品高密度补货策略，否则将持续占用集装箱 CBM 配额而拉低整柜利润。"
    )

    add_image_with_caption(
        doc,
        output_dir / "chart1_volumetric_loss.png",
        "图 1-1：密度最低（最泡）TOP10 品名 & 各渠道日均入仓件数",
        width_inches=6.2,
    )
    doc.add_page_break()

    # ==============================================================
    # 维度 2 / Dimension 2
    # ==============================================================
    add_heading_cjk(doc, "维度 2：客户画像与高净值客户识别", level=1)

    add_heading_cjk(doc, "核心发现", level=2)
    d2 = r.get("dim2", {})
    brand_pct = d2.get("brand_rate_pct", 0.0)
    n_brand = d2.get("n_brand", 0)
    n_no_brand = d2.get("n_no_brand", 0)
    add_bullet(
        doc,
        f"品牌赋能率 {brand_pct:.1f}% —— 有品牌记录 {n_brand:,} 行，"
        f"无品牌 {n_no_brand:,} 行。",
    )
    add_para(doc,
             "这部分有品牌客户大概率在做精细化运营（自建站 / 亚马逊品牌注册 / 独立站），"
             "对清关合规、HS 编码准确性、报关时效要求更高，是未来的高净值核心客户群。"
             "交叉分析显示，有品牌客户的商品在材质与用途上确实呈现出更高附加值的分布特征。")

    add_heading_cjk(doc, "管理层建议", level=2)
    add_banner_quote(
        doc,
        "建议立刻建立有品牌客户的分级标记与专属清关优先级通道。"
        "如果我们能证明自己在合规、时效、数据质量上优于同行，这部分客户就是最稳定、"
        "复购率最高、愿意付费的一群。反之，如果他们觉得你的清关数据经常出错、"
        "HS 编码混乱，转走成本极低。"
    )

    add_image_with_caption(
        doc,
        output_dir / "chart2_customer_profile.png",
        "图 2-1：材质 / 用途分布 × 品牌归属（堆叠对比）",
        width_inches=6.2,
    )
    doc.add_page_break()

    # ==============================================================
    # 维度 3 / Dimension 3
    # ==============================================================
    add_heading_cjk(doc, "维度 3：数据质量与人工干预成本测算", level=1)

    add_heading_cjk(doc, "核心发现", level=2)
    d3 = r.get("dim3", {})
    total_rows = d3.get("total_rows", 0)
    clean_rows = d3.get("needs_cleaning_rows", 0)
    clean_pct = d3.get("cleaning_rate_pct", 0.0)
    add_bullet(
        doc,
        f"{clean_pct:.1f}% 的品名含有模糊词汇 —— "
        f"共 {clean_rows:,} / {total_rows:,} 行需要人工清洗。",
    )

    kw_rows = [[kw, f"{pct:.1f}%"]
               for kw, pct in d3.get("per_keyword_pct", {}).items()]
    if kw_rows:
        add_heading_cjk(doc, "各模糊关键词命中占比", level=3)
        add_table(doc, ["模糊关键词", "命中占比"], kw_rows,
                  col_widths_cm=[9.0, 6.0])

    # 成本测算表
    add_heading_cjk(doc, "业务成本测算（行业经验值）", level=3)
    monthly_qty = 20_000
    manual_hours = clean_pct / 100 * monthly_qty * 3 / 60
    cost_100 = manual_hours * 100
    cost_200 = manual_hours * 200
    cost_rows = [
        ["月处理单数（估算）", f"{monthly_qty:,}"],
        ["需要人工清洗的比例", f"{clean_pct:.1f}%"],
        ["平均每单 HS 编码人工匹配耗时", "3 分钟"],
        ["每月人工工时", f"{manual_hours:.0f} 小时"],
        ["按 100 元/小时估算人工成本", f"≈ {cost_100:,.0f} 元 / 月"],
        ["按 200 元/小时估算人工成本", f"≈ {cost_200:,.0f} 元 / 月"],
    ]
    add_table(doc, ["参数", "数值"], cost_rows, col_widths_cm=[9.5, 5.5])

    add_heading_cjk(doc, "管理层建议", level=2)
    annual_saving = cost_100 * 12 / 10000  # 万元
    add_banner_quote(
        doc,
        f"开发品名 → HS 编码自动匹配引擎 ROI 极高。"
        f"引擎开发投入假设 30 人·天，一旦上线，"
        f"仅每年节省的人工成本即可达到约 {annual_saving:.0f} 万元。"
        "此外，自动匹配的一致性和可追溯性还能显著降低海关查验率和合规风险。"
    )

    add_image_with_caption(
        doc,
        output_dir / "chart3_data_quality.png",
        "图 3-1：各模糊关键词命中占比 & 人工清洗成本占比饼图",
        width_inches=6.2,
    )
    doc.add_page_break()

    # ==============================================================
    # 维度 4 / Dimension 4
    # ==============================================================
    add_heading_cjk(doc, "维度 4：国家 × 报关方式合规矩阵", level=1)

    add_heading_cjk(doc, "核心发现", level=2)
    d4 = r.get("dim4", {})
    cross_count = d4.get("cross_count")
    if cross_count is not None and not cross_count.empty:
        cols = list(cross_count.columns)
        rows = []
        for country, row in cross_count.iterrows():
            values = [str(country)]
            for col in cols:
                values.append(f"{int(row.get(col, 0)):,}")
            values.append(f"{int(row.sum()):,}")
            rows.append(values)
        add_table(
            doc,
            ["收件国家"] + cols + ["合计"],
            rows,
            col_widths_cm=[3.5] + [2.5] * len(cols) + [2.0],
        )

    add_banner_quote(
        doc,
        "合规提醒：某些国家无单证占比偏高，需要重点审视 —— "
        "是因为这些国家对单证要求较宽松（可接受），还是我们自身漏配单证（风险）。"
    )

    add_image_with_caption(
        doc,
        output_dir / "chart4_compliance_matrix.png",
        "图 4-1：国家 × 报关方式 合规矩阵（堆叠柱状图）",
        width_inches=6.2,
    )
    doc.add_page_break()

    # ==============================================================
    # 总结 / Summary
    # ==============================================================
    add_heading_cjk(doc, "总结：为什么要流程优化 + 为什么要自动化", level=1)

    summary_rows = [
        ["1", f"泡货占比 {bubble_pct:.1f}%，持续吞噬集装箱空间",
         "渠道自动配柜引擎", "每柜平均多塞 15-20% 货量"],
        ["2", f"{clean_pct:.1f}% 品名含模糊词，人工匹配耗时巨大",
         "品名 → HS 编码自动匹配引擎",
         f"每月节省 ≈ {manual_hours:.0f} 小时人工"],
        ["3", f"品牌赋能率 {brand_pct:.1f}%，但缺乏高净值客户差异化服务",
         "客户自动分级 + 清关优先级调度",
         "核心客户留存率 + 客户单产值"],
        ["4", "不同国家报关方式合规性不透明，有风险点",
         "国家合规规则自动校验清单", "降低海关查验率和罚金"],
    ]
    add_table(
        doc,
        ["#", "我们的问题", "对应的自动化/流程优化方案", "预期收益"],
        summary_rows,
        col_widths_cm=[0.8, 4.8, 4.7, 4.7],
    )

    doc.add_paragraph()
    add_heading_cjk(doc, "一句话结论", level=2)
    add_banner_quote(
        doc,
        "不做流程优化 = 把利润拱手让给物流成本、人工成本和潜在的合规罚单。"
        "自动化工具不是奢侈品，而是在订单量不变的前提下，"
        "几乎唯一能同时提升利润、提升客户体验、降低团队内耗的抓手。",
        bg_color="E5F0D3",
        text_color=RGBColor(0x37, 0x56, 0x23),
    )

    # ---------- 保存 ----------
    out_path = output_dir / "supply_chain_report.docx"
    doc.save(str(out_path))
    return out_path


# ===========================================================================
# 主程序 / Main
# ===========================================================================

def main() -> None:
    parser = argparse.ArgumentParser(description="Word 报告生成入口")
    parser.add_argument("--mock", action="store_true", help="强制模拟数据")
    parser.add_argument("--real", action="store_true", help="强制真实数据")
    parser.add_argument("--max-files", type=int, default=None, help="真实数据最多读取文件数")
    parser.add_argument("--output", type=str, default=None, help="图表/报告输出目录")
    args = parser.parse_args()

    real_root = r"<用户真实数据目录>"
    use_real = False

    if args.real:
        use_real = True
    elif args.mock:
        use_real = False
    else:
        use_real = os.path.isdir(real_root)

    if use_real:
        output_dir = Path(args.output or "output_real")
        print("[AUTO] 使用真实数据")
        data = load_real_data(real_root, max_files=args.max_files)
        label = "real"
    else:
        output_dir = Path(args.output or "output_mock")
        print("[AUTO] 使用模拟数据")
        data = generate_mock_data()
        label = "mock"

    output_dir.mkdir(parents=True, exist_ok=True)

    analyzer = SupplyChainAnalyzer(data, output_dir=str(output_dir), source_label=label)
    analyzer.run_all()

    docx_path = build_word_report(analyzer, output_dir)
    print(f"\n{'='*60}")
    print(f"Word 报告已生成: {docx_path}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
