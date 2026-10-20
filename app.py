# -*- coding: utf-8 -*-
"""
秋招投递记录 —— 本地网页录入 + Excel 自动写入

运行:  python app.py     （或双击 启动录入.bat）
网页:  http://127.0.0.1:8760
数据:  同目录下 投递记录.xlsx （网页每次提交都会实时写入这个文件）
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import socket
import sys
import threading
import webbrowser
from datetime import date, datetime
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

# 控制台代码页通常是 GBK —— 只放宽错误处理、不改编码，否则中文会变乱码
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(errors="replace")
    except Exception:
        pass

# ---------------------------------------------------------------- 基础配置
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(BASE_DIR, "投递记录.xlsx")
WEB_DIR = os.path.join(BASE_DIR, "web")
PORT = 8760

SHEET = "投递记录"
OPT_SHEET = "选项"
STAT_SHEET = "统计"
TRASH_SHEET = "已删除"      # 删除的记录不销毁，挪到这张表，可一键还原
BACKUP_DIR = os.path.join(BASE_DIR, "backups")
MAX_BACKUPS = 400          # 备份份数上限（每份约 15~100KB）

MAX_ROW = 2000          # 下拉 / 条件格式预覆盖行数
STARTER_ROWS = 60       # 建表时先画好格式的空行数（写入新行时会自动补格式）
MAX_OPT = 60            # 每个下拉列表预留选项数

# 下拉选项（"选项" 工作表里可以直接改，改完网页和 Excel 同步生效）
OPTIONS: dict[str, list[str]] = {
    "渠道": ["BOSS直聘", "猎聘", "智联招聘", "前程无忧", "牛客网", "实习僧",
             "应届生求职网", "公司官网", "内推", "学校就业网", "招聘会", "邮件", "其他"],
    "状态": ["已投递", "简历筛选中", "笔试通知", "笔试完成", "一面通知", "一面完成",
             "二面通知", "二面完成", "HR面通知", "HR面完成", "已发Offer", "已挂", "已放弃"],
    "面试形式": ["线上视频", "线下面试", "电话面试", "AI面试", "待定"],
    "结果": ["待定", "推进中", "通过", "未通过", "已发Offer", "已放弃"],
}

# (标题, 列宽)  —— 顺序即 Excel 列顺序，改这里就同步改整张表
COLUMNS = [
    ("序号", 6),
    ("公司", 20),
    ("岗位", 20),
    ("投递日期", 12),
    ("渠道", 12),
    ("状态", 13),
    ("下一步动作", 22),
    ("截止/跟进时间", 15),
    ("薪资", 14),
    ("工作地点", 13),
    ("备注", 30),
    ("笔试时间", 12),
    ("一面时间", 12),
    ("二面时间", 12),
    ("HR面时间", 12),
    ("面试形式", 12),
    ("结果", 11),
    ("投递天数", 10),
    ("跟进提醒", 18),
]
HEADERS = [c[0] for c in COLUMNS]
COL_IDX = {name: i + 1 for i, (name, _) in enumerate(COLUMNS)}

LIST_FIELDS = {"渠道": "渠道", "状态": "状态", "面试形式": "面试形式", "结果": "结果"}
ALL_DEFAULTS = ["投递日期", "渠道", "状态", "面试形式", "结果"]   # 批量导入时的默认值字段
TEXT_FIELDS = ["公司", "岗位", "下一步动作", "薪资", "工作地点", "备注"]
DATE_FIELDS = ["投递日期", "截止/跟进时间", "笔试时间", "一面时间", "二面时间", "HR面时间"]
EDITABLE = TEXT_FIELDS + DATE_FIELDS + list(LIST_FIELDS)

COL_R = get_column_letter(COL_IDX["投递天数"])   # R
COL_S = get_column_letter(COL_IDX["跟进提醒"])   # S

# ---------------------------------------------------------------- 样式
C_HEADER_BG = "1F4E79"
C_HEADER_FG = "FFFFFF"
C_GROUP_A = "DDEBF7"     # 基础信息
C_GROUP_B = "E2EFDA"     # 跟进
C_GROUP_C = "FFF2CC"     # 面试
C_GROUP_D = "FCE4D6"     # 自动列
C_BORDER = "B4C6E7"
C_ZEBRA = "F7FBFF"

THIN = Side(style="thin", color=C_BORDER)
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

GROUP_OF = {}
for _h in ("序号", "公司", "岗位", "投递日期", "渠道", "薪资", "工作地点"):
    GROUP_OF[_h] = C_GROUP_A
for _h in ("状态", "下一步动作", "截止/跟进时间", "备注"):
    GROUP_OF[_h] = C_GROUP_B
for _h in ("笔试时间", "一面时间", "二面时间", "HR面时间", "面试形式", "结果"):
    GROUP_OF[_h] = C_GROUP_C
for _h in ("投递天数", "跟进提醒"):
    GROUP_OF[_h] = C_GROUP_D


# ---------------------------------------------------------------- 工具
def _to_iso(v) -> str:
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    return str(v).strip()


def _parse_date(s):
    s = (s or "").strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


# ---------------------------------------------------------------- 建表
def style_row(ws, r: int) -> None:
    """给某一行套上边框 / 字体 / 对齐 / 日期格式"""
    for i in range(1, len(COLUMNS) + 1):
        cell = ws.cell(row=r, column=i)
        cell.border = BORDER
        cell.font = Font(name="微软雅黑", size=10)
        cell.alignment = Alignment(vertical="center",
                                   wrap_text=(i == COL_IDX["备注"]))
        if r % 2 == 0:
            cell.fill = PatternFill("solid", fgColor=C_ZEBRA)
    for f in DATE_FIELDS:
        ws.cell(row=r, column=COL_IDX[f]).number_format = "yyyy-mm-dd"
    for f in ("状态", "渠道", "结果", "面试形式", "投递天数", "跟进提醒", "薪资", "工作地点"):
        ws.cell(row=r, column=COL_IDX[f]).alignment = Alignment(
            horizontal="center", vertical="center")


def next_row(ws) -> int:
    """第一条真正空白的行（避开只是被预格式化过的空行）"""
    for r in range(2, ws.max_row + 2):
        if not ws.cell(row=r, column=COL_IDX["公司"]).value and \
           not ws.cell(row=r, column=COL_IDX["岗位"]).value:
            return r
    return ws.max_row + 1


def create_workbook() -> None:
    wb = Workbook()

    # ---------- 投递记录 ----------
    ws = wb.active
    ws.title = SHEET
    for i, (name, width) in enumerate(COLUMNS, start=1):
        c = ws.cell(row=1, column=i, value=name)
        c.font = Font(name="微软雅黑", size=10, bold=True, color=C_HEADER_FG)
        c.fill = PatternFill("solid", fgColor=C_HEADER_BG)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[1].height = 30

    # 表头色带提示（分组）
    for name, color in GROUP_OF.items():
        ws.cell(row=1, column=COL_IDX[name]).fill = PatternFill("solid", fgColor=color)
        ws.cell(row=1, column=COL_IDX[name]).font = Font(
            name="微软雅黑", size=10, bold=True, color="1F4E79")

    for r in range(2, STARTER_ROWS + 2):
        style_row(ws, r)

    ws.freeze_panes = "D2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{MAX_ROW}"

    # ---------- 选项 ----------
    opt = wb.create_sheet(OPT_SHEET)
    opt.column_dimensions["A"].width = 35
    for col, (key, values) in enumerate(OPTIONS.items(), start=1):
        L = get_column_letter(col)
        opt.column_dimensions[L].width = max(
            opt.column_dimensions[L].width or 0, 16)
        h = opt.cell(row=1, column=col, value=key)
        h.font = Font(name="微软雅黑", size=10, bold=True, color=C_HEADER_FG)
        h.fill = PatternFill("solid", fgColor=C_HEADER_BG)
        h.alignment = Alignment(horizontal="center", vertical="center")
        for i, v in enumerate(values, start=2):
            c = opt.cell(row=i, column=col, value=v)
            c.font = Font(name="微软雅黑", size=10)
            c.alignment = Alignment(horizontal="center")
    opt.cell(row=1, column=6, value="公司库（可自行补充，用于网页下拉建议）").font = Font(
        name="微软雅黑", size=10, bold=True, color=C_HEADER_FG)
    opt.cell(row=1, column=6).fill = PatternFill("solid", fgColor=C_HEADER_BG)
    opt.column_dimensions["F"].width = 34

    # ---------- 定义名称（下拉数据源）----------
    for col, key in enumerate(OPTIONS.keys(), start=1):
        L = get_column_letter(col)
        wb.defined_names.add(DefinedName(
            f"{key}选项", attr_text=f"'{OPT_SHEET}'!${L}$2:${L}${MAX_OPT + 1}"))

    # ---------- 数据验证 ----------
    for field, key in LIST_FIELDS.items():
        col = get_column_letter(COL_IDX[field])
        dv = DataValidation(type="list", formula1=f"={key}选项",
                            allow_blank=True, showDropDown=False)
        dv.error = f"请从下拉列表中选择有效的「{field}」"
        dv.errorTitle = "输入无效"
        dv.prompt = f"点击右侧箭头选择{field}"
        dv.promptTitle = field
        ws.add_data_validation(dv)
        dv.add(f"{col}2:{col}{MAX_ROW}")

    # ---------- 条件格式 ----------
    rng = f"A2:{get_column_letter(len(COLUMNS))}{MAX_ROW}"
    F, AF, RF, DF, ORF, PEND = ('$F2="已挂"', '$F2="已放弃"', '$F2="已发Offer"',
                                '$F2="已投递"', '$F2="简历筛选中"',
                                'OR($F2="笔试通知",$F2="一面通知",$F2="二面通知",$F2="HR面通知")')
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[AF], fill=PatternFill("solid", fgColor="F2F2F2"), stopIfTrue=False))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[RF], fill=PatternFill("solid", fgColor="C6EFCE"), stopIfTrue=False))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[DF], fill=PatternFill("solid", fgColor="FFF2CC"), stopIfTrue=False))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[ORF], fill=PatternFill("solid", fgColor="FFE0B2"), stopIfTrue=False))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[PEND], fill=PatternFill("solid", fgColor="E8F4FD"), stopIfTrue=False))
    # 逾期跟进 → 整行浅红
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND(${get_column_letter(COL_IDX["截止/跟进时间"])}2<>"",'
                 f'${get_column_letter(COL_IDX["截止/跟进时间"])}2<TODAY(),'
                 f'$F2<>"已发Offer",$F2<>"已挂",$F2<>"已放弃")'],
        fill=PatternFill("solid", fgColor="FFC7CE"), stopIfTrue=True))

    build_stat_sheet(wb)
    ensure_trash(wb)
    wb.calculation.fullCalcOnLoad = True
    wb.save(XLSX)
    wb.close()


def build_stat_sheet(wb) -> None:
    if STAT_SHEET in wb.sheetnames:
        del wb[STAT_SHEET]
    st = wb.create_sheet(STAT_SHEET)
    st.column_dimensions["A"].width = 4
    st.column_dimensions["B"].width = 18
    st.column_dimensions["C"].width = 12
    st.column_dimensions["D"].width = 4
    st.column_dimensions["E"].width = 18
    st.column_dimensions["F"].width = 12
    st.sheet_view.showGridLines = False

    Q = f"'{SHEET}'!"
    def title(cell, text):
        c = st[cell]
        c.value = text
        c.font = Font(name="微软雅黑", size=12, bold=True, color=C_HEADER_FG)
        c.fill = PatternFill("solid", fgColor=C_HEADER_BG)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BORDER

    def label(cell, text, bold=False, color=None):
        c = st[cell]
        c.value = text
        c.font = Font(name="微软雅黑", size=10, bold=bold, color=color or "000000")
        c.border = BORDER
        c.alignment = Alignment(horizontal="left", vertical="center")

    def num(cell, formula, fmt="0"):
        c = st[cell]
        c.value = formula
        c.font = Font(name="微软雅黑", size=10)
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.number_format = fmt

    st["B2"] = "📊 投递总览（全部自动计算，无需手动改）"
    st["B2"].font = Font(name="微软雅黑", size=13, bold=True, color="1F4E79")
    st.merge_cells("B2:F2")

    # 关键指标
    title("B4", "关键指标")
    title("C4", "数值")
    metrics = [
        ("总投递数", f'=COUNTA({Q}$B$2:$B${MAX_ROW})'),
        ("待跟进（未出结果）", f'=COUNTIFS({Q}$F$2:$F${MAX_ROW},"<>已挂",{Q}$F$2:$F${MAX_ROW},"<>已放弃",{Q}$F$2:$F${MAX_ROW},"<>已发Offer",{Q}$F$2:$F${MAX_ROW},"<>")'),
        ("进入面试环节", f'=COUNTIFS({Q}$F$2:$F${MAX_ROW},"*面*")'),
        ("已发 Offer", f'=COUNTIF({Q}$F$2:$F${MAX_ROW},"已发Offer")'),
        ("已挂 / 已放弃", f'=COUNTIF({Q}$F$2:$F${MAX_ROW},"已挂")+COUNTIF({Q}$F$2:$F${MAX_ROW},"已放弃")'),
        ("近 7 天投递", f'=COUNTIFS({Q}$D$2:$D${MAX_ROW},">="&TODAY()-7)'),
        ("今日投递", f'=COUNTIFS({Q}$D$2:$D${MAX_ROW},TODAY())'),
        ("逾期未跟进", f'=COUNTIFS({Q}$H$2:$H${MAX_ROW},"<"&TODAY(),{Q}$H$2:$H${MAX_ROW},"<>")'),
    ]
    r = 5
    for name, f in metrics:
        label(f"B{r}", name)
        num(f"C{r}", f)
        r += 1
    label(f"B{r}", "Offer 转化率", bold=True)
    num(f"C{r}", f'=IFERROR(COUNTIF({Q}$F$2:$F${MAX_ROW},"已发Offer")/COUNTA({Q}$B$2:$B${MAX_ROW}),0)', "0.0%")

    # 状态分布
    title("E4", "状态分布")
    title("F4", "数量")
    rr = 5
    for s in OPTIONS["状态"]:
        label(f"E{rr}", s)
        num(f"F{rr}", f'=COUNTIF({Q}$F$2:$F${MAX_ROW},"{s}")')
        rr += 1

    # 渠道分布
    title("B17", "渠道分布")
    title("C17", "数量")
    rr = 18
    for s in OPTIONS["渠道"]:
        label(f"B{rr}", s)
        num(f"C{rr}", f'=COUNTIF({Q}$E$2:$E${MAX_ROW},"{s}")')
        rr += 1

    st["E17"] = "💡 提示"
    st["E17"].font = Font(name="微软雅黑", size=11, bold=True, color="1F4E79")
    tips = [
        "· 在「投递记录」表里，状态 / 渠道 / 面试形式 / 结果 都是下拉选择。",
        "· 灰色列「投递天数」「跟进提醒」是自动公式，不用填。",
        "· 整行变红 = 截止/跟进时间已过且还没结果，该催了。",
        "· 「选项」表里可以增删下拉内容，网页表单会同步更新。",
        "· 更推荐用网页录入：双击 启动录入.bat，填完点提交即写入本文件。",
    ]
    rr = 18
    for t in tips:
        c = st.cell(row=rr, column=5, value=t)
        c.font = Font(name="微软雅黑", size=10)
        c.alignment = Alignment(horizontal="left", vertical="center")
        st.merge_cells(start_row=rr, start_column=5, end_row=rr, end_column=6)
        rr += 1


def refresh_formulas(ws, start=2, end=None) -> None:
    """重写自动列公式（序号 / 投递天数 / 跟进提醒）"""
    end = end or ws.max_row
    D = get_column_letter(COL_IDX["投递日期"])
    H = get_column_letter(COL_IDX["截止/跟进时间"])
    for r in range(start, end + 1):
        if not ws.cell(row=r, column=COL_IDX["公司"]).value and \
           not ws.cell(row=r, column=COL_IDX["岗位"]).value:
            ws.cell(row=r, column=COL_IDX["序号"]).value = None
            ws.cell(row=r, column=COL_IDX["投递天数"]).value = None
            ws.cell(row=r, column=COL_IDX["跟进提醒"]).value = None
            continue
        ws.cell(row=r, column=COL_IDX["序号"]).value = "=ROW()-1"
        ws.cell(row=r, column=COL_IDX["投递天数"]).value = (
            f'=IF(${D}{r}="","",TODAY()-${D}{r})')
        ws.cell(row=r, column=COL_IDX["跟进提醒"]).value = (
            f'=IF(${H}{r}="","",'
            f'IF(${H}{r}<TODAY(),"已逾期"&(TODAY()-${H}{r})&"天",'
            f'IF(${H}{r}-TODAY()<=3,"即将到期("&(${H}{r}-TODAY())&"天)","还有"&(${H}{r}-TODAY())&"天")))')


# ---------------------------------------------------------------- 读写
def ensure_workbook() -> None:
    if not os.path.exists(XLSX):
        create_workbook()
        return
    wb = _load()
    if SHEET not in wb.sheetnames:
        wb.close()
        create_workbook()
    wb.close()


def read_options() -> dict:
    """从「选项」表读回下拉内容，让用户在 Excel 里改了也能生效"""
    out = {k: list(v) for k, v in OPTIONS.items()}
    companies: list[str] = []
    try:
        wb = _load(read_only=True, data_only=True)
        if OPT_SHEET in wb.sheetnames:
            opt = wb[OPT_SHEET]
            for col, key in enumerate(OPTIONS.keys(), start=1):
                vals = []
                for r in range(2, MAX_OPT + 2):
                    v = opt.cell(row=r, column=col).value
                    if v is not None and str(v).strip():
                        vals.append(str(v).strip())
                if vals:
                    out[key] = vals
            for r in range(2, 400):
                v = opt.cell(row=r, column=6).value
                if v is not None and str(v).strip():
                    companies.append(str(v).strip())
        wb.close()
    except Exception:
        pass
    out["_公司库"] = companies
    return out


def read_records() -> list[dict]:
    ensure_workbook()
    wb = _load(data_only=False)
    ws = wb[SHEET]
    out = []
    for r in range(2, ws.max_row + 1):
        rec = {"_row": r}
        for name in EDITABLE:
            rec[name] = _to_iso(ws.cell(row=r, column=COL_IDX[name]).value)
        if not rec.get("公司") and not rec.get("岗位"):
            continue
        rec["序号"] = r - 1
        out.append(rec)
    wb.close()
    out.reverse()          # 最新在最上面
    return out


def _apply_row(ws, r: int, data: dict) -> None:
    for name in TEXT_FIELDS:
        v = (data.get(name) or "").strip()
        ws.cell(row=r, column=COL_IDX[name]).value = v or None
    for name in DATE_FIELDS:
        d = _parse_date(data.get(name) or "")
        c = ws.cell(row=r, column=COL_IDX[name])
        c.value = d
        c.number_format = "yyyy-mm-dd"
    for name in LIST_FIELDS:
        v = (data.get(name) or "").strip()
        ws.cell(row=r, column=COL_IDX[name]).value = v or None


# -------------------------------------------------- 批量导入
def _parse_bulk_line(line: str):
    """把一行文本解析成 (公司, 岗位, 日期)。
    支持: '腾讯 后端开发' / '腾讯,后端开发' / '腾讯<Tab>后端开发' / '腾讯 后端开发 2025-09-01'
    """
    s = line.strip().lstrip("-—•*#· \t")
    if not s:
        return "", "", None

    # 行尾可选日期：2025-09-01 / 9-1 / 2025.9.1
    d = None
    m = re.search(r"(\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{1,2}[-/.]\d{1,2})\s*$", s)
    if m:
        raw = m.group(1)
        if len(re.split(r"[-/.]", raw)[0]) == 4:
            d = _parse_date(raw)
        else:
            d = _parse_date(f"{date.today().year}-" + re.sub(r"[-/.]", "-", raw))
        if d:
            s = s[:m.start()].strip()

    for sep in ("\t", ",", "，", ";", "；", "|", "  ", " - ", " / ", "："):
        if sep in s:
            parts = [p.strip() for p in s.split(sep) if p.strip()]
            if len(parts) >= 2:
                return parts[0], " ".join(parts[1:]), d

    parts = s.split()
    if len(parts) >= 2:
        return parts[0], " ".join(parts[1:]), d
    return s, "", d


def _bulk_write(items: list[dict], defaults: dict) -> dict:
    """items: [{公司, 岗位, 可选其它字段}]。一次写入，只备份一次。"""
    parsed, skipped, seen = [], [], set()
    for it in items:
        co = (it.get("公司") or "").strip()
        pos = (it.get("岗位") or "").strip()
        if not co or not pos:
            raw = (it.get("_raw") or f"{co} {pos}").strip()
            skipped.append(raw or "（空行）")
            continue
        if (co, pos) in seen:            # 同一批里重复的只进一次
            continue
        seen.add((co, pos))
        parsed.append((co, pos, it))

    if not parsed:
        return {"added": [], "skipped": skipped, "count": 0}

    ensure_workbook()
    backup_now()
    wb = _load()
    ws = wb[SHEET]
    # 已存在的「公司+岗位」也跳过，避免重复导入
    existing = {(ws.cell(row=r, column=COL_IDX["公司"]).value,
                 ws.cell(row=r, column=COL_IDX["岗位"]).value)
                for r in range(2, ws.max_row + 1)}

    added = []
    for co, pos, it in parsed:
        if (co, pos) in existing:
            skipped.append(f"{co} {pos}（已存在）")
            continue
        r = next_row(ws)
        style_row(ws, r)
        data = dict(defaults)
        for k, v in it.items():          # 逐行覆盖默认值
            if k in COL_IDX and k not in ("序号", "投递天数", "跟进提醒") \
                    and v not in (None, ""):
                data[k] = v
        data["公司"], data["岗位"] = co, pos
        _apply_row(ws, r, data)
        refresh_formulas(ws, r, r)
        existing.add((co, pos))
        added.append(f"{co} · {pos}")

    wb.calculation.fullCalcOnLoad = True
    _safe_save(wb)
    return {"added": added, "skipped": skipped, "count": len(added)}


def bulk_add(text: str, defaults: dict) -> dict:
    """整段粘贴的文本 → 逐行解析后写入"""
    items = []
    for ln in (text or "").splitlines():
        if not ln.strip():
            continue
        co, pos, d = _parse_bulk_line(ln)
        it = {"公司": co, "岗位": pos, "_raw": ln.strip()}
        if d:
            it["投递日期"] = d.isoformat()
        items.append(it)
    return _bulk_write(items, defaults)


def bulk_add_rows(rows: list, defaults: dict) -> dict:
    """结构化行导入：每行可带任意字段（渠道/状态/下一步动作/截止时间…）"""
    return _bulk_write([r for r in rows if isinstance(r, dict)], defaults)


# -------------------------------------------------- 备份 / 回收站（防误删）
def _file_hash(path: str) -> str:
    with open(path, "rb") as f:
        return hashlib.sha1(f.read()).hexdigest()


def _backup_names() -> list[str]:
    if not os.path.isdir(BACKUP_DIR):
        return []
    return sorted(f for f in os.listdir(BACKUP_DIR)
                  if f.startswith("投递记录_") and f.endswith(".xlsx"))


def backup_now(force: bool = False) -> str | None:
    """把当前 Excel 快照进 backups/。内容与上一份相同则跳过。"""
    if not os.path.exists(XLSX):
        return None
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        names = _backup_names()
        if names and not force:
            if _file_hash(os.path.join(BACKUP_DIR, names[-1])) == _file_hash(XLSX):
                return None
        ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
        dst = os.path.join(BACKUP_DIR, f"投递记录_{ts}.xlsx")
        shutil.copy2(XLSX, dst)
        for old in _backup_names()[:-MAX_BACKUPS]:
            try:
                os.remove(os.path.join(BACKUP_DIR, old))
            except OSError:
                pass
        return dst
    except Exception:
        return None          # 备份失败绝不阻断用户操作


def list_backups(limit: int = 40) -> list[dict]:
    out = []
    for name in reversed(_backup_names()):
        if len(out) >= limit:
            break
        p = os.path.join(BACKUP_DIR, name)
        try:
            st = os.stat(p)
        except OSError:
            continue
        stamp = name[len("投递记录_"):-len(".xlsx")]
        out.append({
            "name": name,
            "size": st.st_size,
            "time": datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
        })
    return out


def restore_backup(name: str) -> str:
    src = os.path.join(BACKUP_DIR, os.path.basename(name))
    if not os.path.isfile(src):
        raise KeyError("备份文件不存在")
    backup_now(force=True)          # 恢复前先存当前状态，可再退回来
    try:
        shutil.copy2(src, XLSX)
    except PermissionError:
        raise RuntimeError(BUSY_MSG)
    return os.path.basename(name)


def ensure_trash(wb):
    """「已删除」表：删除的记录挪到这里，不销毁"""
    if TRASH_SHEET not in wb.sheetnames:
        ts = wb.create_sheet(TRASH_SHEET)
        heads = ["删除时间"] + HEADERS
        for i, h in enumerate(heads, start=1):
            c = ts.cell(row=1, column=i, value=h)
            c.font = Font(name="微软雅黑", size=10, bold=True, color=C_HEADER_FG)
            c.fill = PatternFill("solid", fgColor="9C6B4E")
            c.alignment = Alignment(horizontal="center", vertical="center")
        ts.column_dimensions["A"].width = 20
        for i, (_h, wdt) in enumerate(COLUMNS, start=2):
            ts.column_dimensions[get_column_letter(i)].width = wdt
        ts.freeze_panes = "B2"
    return wb[TRASH_SHEET]


def read_trash() -> list[dict]:
    ensure_workbook()
    wb = _load(data_only=False)
    if TRASH_SHEET not in wb.sheetnames:
        wb.close()
        return []
    ts = wb[TRASH_SHEET]
    out = []
    for r in range(2, ts.max_row + 1):
        rec = {"_trow": r, "删除时间": _to_iso(ts.cell(row=r, column=1).value)}
        for i, name in enumerate(HEADERS, start=1):
            if name in ("序号", "投递天数", "跟进提醒"):
                continue
            rec[name] = _to_iso(ts.cell(row=r, column=i + 1).value)
        if not rec.get("公司") and not rec.get("岗位"):
            continue
        out.append(rec)
    wb.close()
    out.reverse()
    return out


def restore_trash(trow: int) -> int:
    ensure_workbook()
    backup_now()
    wb = _load()
    if TRASH_SHEET not in wb.sheetnames:
        wb.close()
        raise KeyError("回收站是空的")
    ts = wb[TRASH_SHEET]
    if trow < 2 or trow > ts.max_row:
        wb.close()
        raise KeyError("这条记录已不存在，请刷新")
    vals = {name: ts.cell(row=trow, column=i + 1).value
            for i, name in enumerate(HEADERS, start=1)}
    ws = wb[SHEET]
    r = next_row(ws)
    style_row(ws, r)
    for name in HEADERS:
        if name in ("序号", "投递天数", "跟进提醒"):
            continue
        c = ws.cell(row=r, column=COL_IDX[name])
        c.value = vals.get(name)
        if name in DATE_FIELDS:
            c.number_format = "yyyy-mm-dd"
    refresh_formulas(ws, r, r)
    ts.delete_rows(trow, 1)
    wb.calculation.fullCalcOnLoad = True
    _safe_save(wb)
    return r - 1


def add_record(data: dict) -> int:
    ensure_workbook()
    backup_now()
    wb = _load()
    ws = wb[SHEET]
    r = next_row(ws)
    style_row(ws, r)
    _apply_row(ws, r, data)
    refresh_formulas(ws, r, r)
    wb.calculation.fullCalcOnLoad = True
    _safe_save(wb)
    return r - 1


def update_record(row: int, data: dict) -> None:
    ensure_workbook()
    backup_now()
    wb = _load()
    ws = wb[SHEET]
    if row < 2 or row > ws.max_row:
        wb.close()
        raise KeyError("记录不存在，请刷新页面")
    style_row(ws, row)
    _apply_row(ws, row, data)
    refresh_formulas(ws, row, row)
    wb.calculation.fullCalcOnLoad = True
    _safe_save(wb)


def delete_record(row: int) -> None:
    ensure_workbook()
    backup_now()
    wb = _load()
    ws = wb[SHEET]
    if row < 2 or row > ws.max_row:
        wb.close()
        raise KeyError("记录不存在，请刷新页面")

    # 关键：不销毁数据。整行先复制进「已删除」表，随时能一键还原。
    ts = ensure_trash(wb)
    tr = ts.max_row + 1
    ts.cell(row=tr, column=1, value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    for i in range(1, len(COLUMNS) + 1):
        ts.cell(row=tr, column=i + 1, value=ws.cell(row=row, column=i).value)

    ws.delete_rows(row, 1)
    refresh_formulas(ws)                      # 移动后公式错位，整体重写
    wb.calculation.fullCalcOnLoad = True
    _safe_save(wb)


def _safe_save(wb) -> None:
    try:
        wb.save(XLSX)
    except PermissionError:
        wb.close()
        raise RuntimeError(BUSY_MSG)
    wb.close()


BUSY_MSG = ("写入失败：投递记录.xlsx 正被 Excel 打开着。\n"
            "请在 Excel 里关掉这个文件，再点一次提交即可。")


def _load(**kw):
    """打开工作簿；被 Excel 独占时给出人话提示而不是 Errno 13"""
    try:
        return load_workbook(XLSX, **kw)
    except PermissionError:
        raise RuntimeError(BUSY_MSG)


def compute_stats(records: list[dict]) -> dict:
    total = len(records)
    offer = sum(1 for r in records if r.get("状态") == "已发Offer")
    dead = sum(1 for r in records if r.get("状态") in ("已挂", "已放弃"))
    interview = sum(1 for r in records if "面" in (r.get("状态") or ""))
    today = date.today()
    week = 0
    overdue = 0
    for r in records:
        d = _parse_date(r.get("投递日期"))
        if d and (today - d).days <= 7:
            week += 1
        h = _parse_date(r.get("截止/跟进时间"))
        if h and h < today and r.get("状态") not in ("已发Offer", "已挂", "已放弃"):
            overdue += 1
    by_status = {}
    for r in records:
        s = r.get("状态") or "（未填）"
        by_status[s] = by_status.get(s, 0) + 1
    return {
        "total": total, "offer": offer, "dead": dead,
        "interview": interview, "week": week, "overdue": overdue,
        "pending": total - offer - dead,
        "rate": round(offer / total * 100, 1) if total else 0.0,
        "by_status": by_status,
    }


# ---------------------------------------------------------------- HTTP
class Handler(BaseHTTPRequestHandler):
    server_version = "QiuzhaoTracker/1.0"

    def log_message(self, fmt, *args):
        pass

    # ---------- helpers ----------
    def _json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _err(self, msg, code=400):
        self._json({"ok": False, "error": msg}, code)

    def _body(self) -> dict:
        n = int(self.headers.get("Content-Length") or 0)
        if not n:
            return {}
        raw = self.rfile.read(n).decode("utf-8")
        return json.loads(raw or "{}")

    def _static(self, path: str):
        if path in ("/", ""):
            path = "/index.html"
        rel = os.path.normpath(path.lstrip("/")).replace("\\", "/")
        if rel.startswith(".."):
            self._err("bad path", 403)
            return
        full = os.path.join(WEB_DIR, rel)
        if not os.path.isfile(full):
            self._err("not found", 404)
            return
        ctype = {
            ".html": "text/html; charset=utf-8",
            ".css": "text/css; charset=utf-8",
            ".js": "application/javascript; charset=utf-8",
            ".svg": "image/svg+xml",
            ".ico": "image/x-icon",
        }.get(os.path.splitext(full)[1].lower(), "application/octet-stream")
        with open(full, "rb") as f:
            data = f.read()
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    # ---------- routes ----------
    def do_GET(self):
        u = urlparse(self.path)
        if u.path == "/api/data":
            try:
                recs = read_records()
                self._json({"ok": True, "records": recs,
                            "options": read_options(),
                            "stats": compute_stats(recs),
                            "backups": list_backups(30),
                            "trash": read_trash(),
                            "file": XLSX})
            except Exception as e:
                self._err(f"读取 Excel 失败：{e}", 500)
            return
        if u.path == "/api/ping":
            self._json({"ok": True, "app": "qiuzhao"})
            return
        self._static(u.path)

    def do_POST(self):
        u = urlparse(self.path)
        try:
            data = self._body()
        except Exception as e:
            self._err(f"请求格式错误：{e}")
            return

        if u.path == "/api/add":
            if not (data.get("公司") or "").strip() or not (data.get("岗位") or "").strip():
                self._err("「公司」和「岗位」必须填写")
                return
            try:
                seq = add_record(data)
                self._json({"ok": True, "序号": seq,
                            "message": f"已写入 Excel 第 {seq} 条：{data['公司']} · {data['岗位']}"})
            except Exception as e:
                self._err(str(e), 500)
            return

        if u.path == "/api/update":
            row = int(data.get("_row") or 0)
            if not (data.get("公司") or "").strip() or not (data.get("岗位") or "").strip():
                self._err("「公司」和「岗位」必须填写")
                return
            try:
                cur = read_records()
                match = next((r for r in cur if r["_row"] == row), None)
                if match and data.get("_expect") and \
                        (match.get("公司"), match.get("岗位")) != tuple(data["_expect"]):
                    self._err("该行内容已被改动，请点「刷新」后重试", 409)
                    return
                update_record(row, data)
                self._json({"ok": True, "message": "已更新"})
            except Exception as e:
                self._err(str(e), 500)
            return

        if u.path == "/api/delete":
            row = int(data.get("_row") or 0)
            try:
                delete_record(row)
                self._json({"ok": True,
                            "message": "已删除（可在「回收站」里一键还原）"})
            except Exception as e:
                self._err(str(e), 500)
            return

        if u.path == "/api/bulk":
            defaults = {k: (data.get(k) or "").strip() for k in ALL_DEFAULTS}
            if not defaults.get("投递日期"):
                defaults["投递日期"] = date.today().isoformat()
            try:
                if isinstance(data.get("rows"), list):
                    res = bulk_add_rows(data["rows"], defaults)
                else:
                    res = bulk_add(data.get("text") or "", defaults)
                msg = f"成功导入 {res['count']} 条"
                if res["skipped"]:
                    msg += f"，跳过 {len(res['skipped'])} 条"
                self._json({"ok": True, "message": msg, **res})
            except Exception as e:
                self._err(str(e), 500)
            return

        if u.path == "/api/backup":
            try:
                p = backup_now(force=True)
                self._json({"ok": True,
                            "message": "已备份" if p else "备份未生成",
                            "name": os.path.basename(p) if p else None})
            except Exception as e:
                self._err(str(e), 500)
            return

        if u.path == "/api/restore":
            try:
                name = restore_backup(data.get("name") or "")
                self._json({"ok": True, "message": f"已恢复到 {name}"})
            except Exception as e:
                self._err(str(e), 500)
            return

        if u.path == "/api/trash/restore":
            try:
                seq = restore_trash(int(data.get("_trow") or 0))
                self._json({"ok": True, "message": f"已还原，现在是第 {seq} 条"})
            except Exception as e:
                self._err(str(e), 500)
            return

        if u.path == "/api/rebuild":
            try:
                if os.path.exists(XLSX):
                    backup_now(force=True)                      # 先存快照
                    os.replace(XLSX, XLSX + ".bak")             # 再留一份 .bak
                create_workbook()
                self._json({"ok": True,
                            "message": "已重建空表（原文件已存入 backups\\ 并留底为 投递记录.xlsx.bak）"})
            except Exception as e:
                self._err(str(e), 500)
            return

        self._err("not found", 404)


def pick_port(start: int) -> int:
    for p in range(start, start + 20):
        with socket.socket() as s:
            if s.connect_ex(("127.0.0.1", p)) != 0:
                return p
    return start


def already_running() -> bool:
    """同目录已在跑 → 直接开浏览器，不重复启动"""
    import urllib.request
    try:
        with urllib.request.urlopen(
                f"http://127.0.0.1:{PORT}/api/ping", timeout=1.2) as r:
            return json.loads(r.read().decode("utf-8")).get("app") == "qiuzhao"
    except Exception:
        return False


def main():
    if already_running():
        webbrowser.open(f"http://127.0.0.1:{PORT}/")
        print("服务已在运行，已为你打开浏览器。")
        return
    ensure_workbook()
    port = pick_port(PORT)
    url = f"http://127.0.0.1:{port}/"
    srv = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    print("=" * 58)
    print("  秋招投递记录 · 本地录入服务已启动")
    print(f"  网页地址: {url}")
    print(f"  数据文件: {XLSX}")
    print("  保持本窗口开着；关掉窗口即停止服务（数据已存进 Excel）。")
    print("=" * 58)
    if not os.environ.get("QP_NOBROWSER"):
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止。")
    finally:
        srv.server_close()


if __name__ == "__main__":
    main()
