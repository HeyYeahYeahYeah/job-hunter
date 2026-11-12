#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""job-hunter 分析脚本：读取 jobs_raw.json -> 生成 Markdown 分析报告 + Excel 明细表。

专为「大专对口岗位 → 课程标准反推」设计，核心能力：
1. 学历过滤（默认只要大专）
2. 能力提取与聚类（岗位技能词 → 能力大类 → 频次排名）
3. 课程映射建议（能力大类 → 建议课程/实训项目）
"""
import json
import sys
import os
from collections import Counter
from datetime import datetime

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("缺少 openpyxl，请先 pip install openpyxl")
    sys.exit(1)


# ============ 能力词典 ============
# 技能关键词（小写匹配）→ (能力大类, 建议课程/实训)
SKILL_MAP = {
    # 控制类
    "plc": ("PLC/工控编程", "《PLC编程与应用》《电气控制技术》"),
    "dcs": ("DCS/集散控制", "《过程控制与集散系统》《DCS组态与维护》"),
    "wincc": ("组态/HMI", "《组态软件与人机界面》《HMI触摸屏组态》"),
    "hmi": ("组态/HMI", "《组态软件与人机界面》《HMI触摸屏组态》"),
    "modbus": ("工业通信", "《工业网络与通信》《现场总线技术》"),
    "profinet": ("工业通信", "《工业网络与通信》《现场总线技术》"),
    "ethercat": ("工业通信", "《工业网络与通信》《现场总线技术》"),
    "mqtt": ("工业通信/物联网", "《物联网通信技术》《工业网络与通信》"),
    "opc": ("工业通信/数据采集", "《工业网络与通信》《SCADA与数据采集》"),
    "通讯协议": ("工业通信", "《工业网络与通信》《现场总线技术》"),
    "伺服": ("运动控制", "《运动控制系统》《伺服与步进驱动》"),
    "步进": ("运动控制", "《运动控制系统》《伺服与步进驱动》"),
    "变频器": ("变频/拖动", "《电力拖动与变频调速》《电机与拖动》"),
    "电机": ("电机与拖动", "《电机与电气控制》《电机与拖动》"),
    "驱动": ("电机与拖动", "《电机与电气控制》《电机与拖动》"),
    # 电气类
    "电气": ("电气识图/电气控制", "《电气控制技术》《电工基础》"),
    "电工": ("电气识图/电气控制", "《电工基础》《电气安装与维修》"),
    "电气原理图": ("电气识图/电气控制", "《电气控制技术》《电工基础》"),
    "接线": ("电气识图/电气控制", "《电气安装与维修》《电工实训》"),
    "控制柜": ("电气识图/电气控制", "《电气控制技术》《电气安装》"),
    "配电": ("电气识图/电气控制", "《电气控制技术》《供配电技术》"),
    "低压": ("电气识图/电气控制", "《电工基础》《低压电器》"),
    "eplan": ("电气制图", "《EPLAN电气制图》"),
    "cad": ("工程制图", "《AutoCAD电气制图》《机械制图》"),
    "solidworks": ("工程制图/三维", "《SolidWorks三维设计》"),
    # 嵌入式/编程
    "单片机": ("单片机/嵌入式", "《单片机原理与应用》《嵌入式系统》"),
    "嵌入式": ("单片机/嵌入式", "《单片机原理与应用》《嵌入式系统》"),
    "stm32": ("单片机/嵌入式", "《单片机原理与应用》《嵌入式系统》"),
    "esp32": ("单片机/嵌入式", "《单片机原理与应用》《嵌入式系统》"),
    "c++": ("C/C++编程", "《C/C++程序设计》"),
    "c语言": ("C/C++编程", "《C/C++程序设计》"),
    "python": ("Python编程", "《Python程序设计》"),
    "java": ("Java编程", "《Java程序设计》"),
    "arm": ("单片机/嵌入式", "《嵌入式系统》《ARM体系结构》"),
    "pcb": ("电子电路", "《电子技术基础》《PCB设计与制作》"),
    "数字电路": ("电子电路", "《电子技术基础》《数字电子技术》"),
    "模拟电子": ("电子电路", "《电子技术基础》《模拟电子技术》"),
    # 仪表/传感器
    "传感器": ("传感器与检测", "《传感器与检测技术》"),
    "仪表": ("仪表/检测", "《工业仪表与检测》《自动化仪表》"),
    "变送器": ("仪表/检测", "《工业仪表与检测》《自动化仪表》"),
    "流量计": ("仪表/检测", "《工业仪表与检测》《自动化仪表》"),
    "压力表": ("仪表/检测", "《工业仪表与检测》《自动化仪表》"),
    "检测": ("传感器与检测", "《传感器与检测技术》"),
    # 机器人/视觉
    "机器人": ("工业机器人", "《工业机器人操作与编程》《机器人工作站集成》"),
    "机器视觉": ("机器视觉", "《机器视觉技术》《图像处理》"),
    "halcon": ("机器视觉", "《机器视觉技术》《图像处理》"),
    "visionpro": ("机器视觉", "《机器视觉技术》《图像处理》"),
    "视觉": ("机器视觉", "《机器视觉技术》《图像处理》"),
    "ros": ("机器人操作系统", "《ROS机器人开发》"),
    # 自动化系统
    "scada": ("SCADA/上位机", "《SCADA系统与上位机》《组态软件》"),
    "mes": ("MES/数字化", "《智能制造与MES系统》《数字化工厂》"),
    "组态": ("组态/HMI", "《组态软件与人机界面》《HMI触摸屏组态》"),
    "上位机": ("SCADA/上位机", "《SCADA系统与上位机》"),
    "自控系统": ("过程控制", "《过程控制与自动化仪表》"),
    "自动化": ("自动化基础", "《自动控制原理》《自动化概论》"),
    "智能工厂": ("数字化工厂", "《智能制造与MES系统》《数字化工厂》"),
    "it/ot": ("数字化工厂", "《智能制造与MES系统》《工业互联网》"),
    "工业物联网": ("工业物联网", "《物联网技术》《工业互联网》"),
    # 算法/仿真
    "pid": ("自动控制原理", "《自动控制原理》《过程控制》"),
    "matlab": ("系统仿真", "《MATLAB与系统仿真》"),
    "simulink": ("系统仿真", "《MATLAB与系统仿真》"),
    "slam": ("机器视觉/算法", "《机器视觉技术》《智能控制》"),
    "算法": ("算法基础", "《算法与数据结构》《智能控制》"),
    "深度学习": ("AI基础", "《人工智能基础》《机器学习》"),
    # 其他
    "linux": ("操作系统", "《Linux操作系统》"),
    "数据库": ("数据库", "《数据库应用技术》"),
    "sql": ("数据库", "《数据库应用技术》"),
    "安全": ("安全规范", "《安全生产与电工安全》"),
    "开关": ("电气元件", "《电工基础》《低压电器》"),
    "西门子": ("PLC/工控编程", "《PLC编程与应用》(西门子S7系列)"),
    "三菱": ("PLC/工控编程", "《PLC编程与应用》(三菱系列)"),
    "欧姆龙": ("PLC/工控编程", "《PLC编程与应用》(欧姆龙系列)"),
}


def extract_abilities(job):
    """从岗位的 skills 字段提取能力，返回 [(能力大类, 建议课程), ...] 列表（去重）。"""
    text = job.get("skills", "") or ""
    found = {}  # 能力大类 -> 建议课程
    tl = text.lower()
    for kw, (cat, course) in SKILL_MAP.items():
        if kw in tl:
            if cat not in found:
                found[cat] = course
    return list(found.items())


def load_jobs(path):
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, dict):
        return data.get("jobs", [])
    return data


def filter_by_education(jobs, edu="大专"):
    """过滤学历：默认只要大专可投岗位。
    匹配规则：学历字段包含'大专'（含'大专可投'）或'学历不限'（大专学历必然可投）。"""
    result = []
    for j in jobs:
        e = j.get("education") or ""
        if edu in e or "学历不限" in e:
            result.append(j)
    return result


def parse_salary(s):
    if not s:
        return None
    s = str(s).replace(",", "")
    if any(k in s for k in ["元/天", "元/时", "元/小时", "/天", "/时", "/小时", "日", "实习", "面议"]):
        return None
    import re
    nums = re.findall(r"(\d+(?:\.\d+)?)", s)
    if not nums:
        return None
    vals = [float(x) for x in nums]
    if "万" in s:
        vals = [v * 10 for v in vals]
    elif "k" in s.lower() or "K" in s:
        pass
    else:
        if max(vals) > 1000:
            vals = [v / 1000 for v in vals]
    if len(vals) == 1:
        return (vals[0], vals[0])
    return (min(vals), max(vals))


def fmt_salary(s):
    if not s:
        return ""
    s = str(s)
    if any(k in s for k in ["元/天", "元/时", "元/小时", "/天", "/时", "/小时", "实习"]):
        return s
    if "面议" in s or "遮挡" in s:
        return "面议"
    r = parse_salary(s)
    if not r:
        return s
    lo, hi = r
    return f"{lo:.1f}K" if lo == hi else f"{lo:.1f}-{hi:.1f}K"


def salary_band(s):
    r = parse_salary(s)
    if not r:
        return "面议/未公开"
    mid = (r[0] + r[1]) / 2
    if mid < 8:
        return "<8K"
    if mid < 15:
        return "8-15K"
    if mid < 25:
        return "15-25K"
    return "25K+"


def build_md(jobs, major, edu_filter):
    n = len(jobs)
    if n == 0:
        return f"# {major}专业（{edu_filter}）岗位能力分析报告\n\n未抓取到符合条件的岗位数据。\n"

    cats = [j.get("category", "其他") for j in jobs]
    cities = [j.get("city", "") for j in jobs]
    inds = [j.get("industry", "") for j in jobs]
    exps = [j.get("experience", "") for j in jobs]

    cat_counter = Counter(cats)
    city_counter = Counter(c for c in cities if c)
    ind_counter = Counter(i for i in inds if i)
    exp_counter = Counter(e for e in exps if e)

    # 能力聚类统计
    ability_counter = Counter()
    ability_course = {}  # 能力大类 -> 建议课程
    job_ability_map = []  # (岗位名, [能力大类...])
    for j in jobs:
        abilities = extract_abilities(j)
        cats_this = [a[0] for a in abilities]
        job_ability_map.append((j.get("title", ""), cats_this))
        for cat, course in abilities:
            ability_counter[cat] += 1
            ability_course[cat] = course

    lines = []
    lines.append(f"# {major}专业（{edu_filter}）岗位能力分析报告\n")
    lines.append(f"> 数据来源：智联招聘等主流平台（在招岗位）")
    lines.append(f"> 学历范围：仅 {edu_filter} 可投")
    lines.append(f"> 岗位样本：{n} 个")
    lines.append(f"> 生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}")
    lines.append(f"> 用途：为课程标准修订提供岗位能力依据\n")

    # 一、岗位分类
    lines.append("## 一、岗位分类\n")
    lines.append("| 岗位类别 | 数量 | 占比 |")
    lines.append("|---|---|---|")
    for k, v in cat_counter.most_common():
        lines.append(f"| {k} | {v} | {v/n*100:.1f}% |")
    lines.append("")

    # 二、能力需求分析（核心）
    lines.append("## 二、岗位能力需求分析（核心）\n")
    lines.append("### 2.1 能力大类频次排名\n")
    lines.append("| 能力大类 | 需求频次 | 覆盖岗位占比 |")
    lines.append("|---|---|---|")
    for k, v in ability_counter.most_common(20):
        lines.append(f"| {k} | {v} | {v/n*100:.1f}% |")
    lines.append("")

    lines.append("### 2.2 能力 → 建议课程/实训映射\n")
    lines.append("下表为能力大类对应的建议课程与实训项目，可直接作为课程标准修订的参考。\n")
    lines.append("| 能力大类 | 需求频次 | 建议课程/实训 |")
    lines.append("|---|---|---|")
    for k, v in ability_counter.most_common(20):
        lines.append(f"| {k} | {v} | {ability_course.get(k, '—')} |")
    lines.append("")

    lines.append("### 2.3 岗位 × 能力对照\n")
    lines.append("| 岗位 | 所需能力 |")
    lines.append("|---|---|")
    for title, cats_this in job_ability_map:
        lines.append(f"| {title} | {', '.join(cats_this) if cats_this else '（未匹配到明确能力词）'} |")
    lines.append("")

    # 三、经验要求
    lines.append("## 三、经验要求\n")
    for k, v in exp_counter.most_common():
        lines.append(f"- {k}：{v} 个（{v/n*100:.1f}%）")
    lines.append("")

    # 四、行业分布
    lines.append("## 四、行业分布\n")
    for k, v in ind_counter.most_common(10):
        lines.append(f"- {k}：{v} 个（{v/n*100:.1f}%）")
    lines.append("")

    # 五、地域分布
    lines.append("## 五、地域分布\n")
    for k, v in city_counter.most_common(10):
        lines.append(f"- {k}：{v} 个（{v/n*100:.1f}%）")
    lines.append("")

    # 六、课程标准修订建议
    lines.append("## 六、课程标准修订建议\n")
    lines.append("结合岗位能力需求，课程标准修订建议如下：\n")
    top_abilities = ability_counter.most_common(8)
    if top_abilities:
        lines.append("**① 应强化的核心能力模块**（按岗位需求频次排序）：\n")
        for k, v in top_abilities:
            lines.append(f"- {k}（{v} 个岗位需求）— {ability_course.get(k, '')}")
        lines.append("")
    lines.append("**② 修订原则**：")
    lines.append("1. 将上述高频能力作为课程目标与考核要点，确保「课程内容 → 岗位能力」一一对应。")
    lines.append("2. 强化实训/项目化教学，把岗位真实任务（如 PLC 控制柜接线调试、HMI 组态、仪表安装）转化为实训项目。")
    lines.append("3. 对低需求/已淘汰内容做减法，避免课程冗余。")
    lines.append("")

    # 七、岗位明细
    lines.append("## 七、岗位明细\n")
    lines.append(f"共 {n} 个岗位，完整明细见《{major}岗位明细.xlsx》。\n")
    lines.append("| 岗位 | 公司 | 城市 | 经验 | 岗位类别 |")
    lines.append("|---|---|---|---|---|")
    for j in jobs[:30]:
        lines.append(f"| {j.get('title','')} | {j.get('company','')} | {j.get('city','')} | {j.get('experience','')} | {j.get('category','')} |")
    if n > 30:
        lines.append(f"\n（仅展示前 30 条，其余见 Excel）")

    return "\n".join(lines)


def build_xlsx(jobs, major, out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "岗位明细"

    headers = ["序号", "岗位名称", "公司", "薪资", "城市", "经验要求", "学历要求", "岗位类别", "行业", "所需能力", "来源", "链接"]
    ws.append(headers)

    header_fill = PatternFill("solid", fgColor="2F5597")
    header_font = Font(bold=True, color="FFFFFF")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    for i, j in enumerate(jobs, 1):
        abilities = extract_abilities(j)
        ability_str = "、".join([a[0] for a in abilities])
        row = [
            i,
            j.get("title", ""),
            j.get("company", ""),
            fmt_salary(j.get("salary", "")),
            j.get("city", ""),
            j.get("experience", ""),
            j.get("education", ""),
            j.get("category", ""),
            j.get("industry", ""),
            ability_str,
            j.get("source", ""),
            j.get("url", ""),
        ]
        ws.append(row)
        for c in range(1, len(headers) + 1):
            ws.cell(row=i + 1, column=c).border = border

    widths = [6, 24, 22, 12, 12, 12, 10, 16, 16, 28, 10, 30]
    for idx, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = w

    ws.freeze_panes = "A2"
    wb.save(out_path)
    return out_path


def main():
    if len(sys.argv) < 2:
        print("用法: python analyze.py <jobs_raw.json> [专业名] [输出目录] [学历:默认大专]")
        sys.exit(1)

    raw = sys.argv[1]
    major = sys.argv[2] if len(sys.argv) > 2 else "专业"
    out_dir = sys.argv[3] if len(sys.argv) > 3 else "."
    edu_filter = sys.argv[4] if len(sys.argv) > 4 else "大专"

    if not os.path.exists(raw):
        print(f"找不到 {raw}")
        sys.exit(1)

    jobs = load_jobs(raw)
    jobs = filter_by_education(jobs, edu_filter)
    md_text = build_md(jobs, major, edu_filter)

    md_path = os.path.join(out_dir, f"{major}岗位分析报告.md")
    xlsx_path = os.path.join(out_dir, f"{major}岗位明细.xlsx")

    with open(md_path, "w", encoding="utf-8") as f:
        f.write(md_text)

    build_xlsx(jobs, major, xlsx_path)

    print(f"总岗位数: {len(load_jobs(raw))}，{edu_filter}岗位数: {len(jobs)}")
    print(f"已生成: {md_path}")
    print(f"已生成: {xlsx_path}")


if __name__ == "__main__":
    main()
