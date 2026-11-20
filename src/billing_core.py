#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
需求2（升级版）：事件对齐 + 分段计费 + 输出报表（同一程序，分两个阶段）

Stage 1：把“申请表(事件流)”对齐成“事件时间线”，再切分成“计费区间段”，生成 billing_detail（按段）
Stage 2：按“中心/课题组”汇总，并输出 Excel（汇总 + 分类型明细 多sheet）

关键规则（与你的业务一致）：
- 新增事件时间点 = 计费开始时间
- 腾退事件时间点 = 计费截止时间
- 同一 (所属中心, 所属课题组, 资源类型, 编号) 形成一条时间线
- 实验室允许多次腾退：每条腾退的"实验室面积"视为本次减少的面积（你会人工填对）
- 费用口径：统一为闭区间（含首尾两日），days = (end - start).days + 1

自测方法（示例）：
- 测试用例：start=2023-10-27, end=2026-01-12
- 闭区间天数 = (end-start).days + 1 = (2026-01-12 - 2023-10-27).days + 1
- Stage1 输出的使用天数应等于该闭区间天数
- 跨月拆分后逐月天数之和应等于该闭区间天数
"""

import re
import math
from dataclasses import dataclass, field
from datetime import datetime, date, timedelta
from typing import Dict, List, Optional, Tuple, Any

import pandas as pd


# ==================== 路径（放前面，方便你改） ====================
APPLICATION_FILE = 'examples/input.xlsx'
OUTPUT_REPORT_FILE = 'outputs/report.xlsx'

# ==================== 可配置参数 ====================
ELECTRICITY_UNIT_PRICE = 1.0  # 你表里“电费单价1元”，接口保留（目前电费字段已是月度电费）
INCLUDE_START_DAY = True      # 固定为含首日（闭区间）：days = (end - start).days + 1
DAY_DIVISOR = 30.0            # 月度折算天数
NEG_TOL = 1e-6                # 允许的负误差（浮点）


# ==================== 列名别名映射（容错） ====================
ALIASES = {
    # 通用
    "序号": ["序号", "编号序号", "No", "NO"],
    "申请日期": ["申请日期", "申请时间", "日期"],
    "计费开始时间": ["计费开始时间", "开始时间", "计费开始", "起始时间"],
    "计费截止时间": ["计费截止时间", "截止时间", "计费截止", "结束时间"],
    "所属中心": ["所属中心", "中心", "所属中心/平台"],
    "所属课题组": ["所属课题组", "课题组", "团队", "使用团队", "正在使用团队"],
    "经办人": ["经办人", "经办", "联系人"],
    "地址": ["地址", "园区", "地点"],
    "楼层": ["楼层", "层", "所在楼层"],
    "新增/腾退": ["新增/腾退", "新增腾退", "变更类型", "类型"],

    # 办公室
    "办公室名称": ["办公室名称", "房间名称", "办公室名"],
    "办公室编号": ["办公室编号", "办公室号", "房间编号", "房间号"],
    "办公室面积": ["办公室面积", "面积(办公室)", "房间面积"],
    "办公室公摊率": ["办公室公摊率", "公摊率(办公室)", "办公室公摊"],
    "办公室月度单价": ["办公室月度单价", "办公室单价", "办公室月单价"],

    # 卡位
    "卡位名称": ["卡位名称", "卡座名称", "工位名称"],
    "卡位编号": ["卡位编号", "卡座编号", "工位编号"],
    "卡位数量": ["卡位数量", "工位数量", "卡座数量", "数量"],
    "卡位面积": ["卡位面积", "面积(卡位)"],
    "卡位公摊率": ["卡位公摊率", "公摊率(卡位)", "卡位公摊"],
    "卡位使用人": ["卡位使用人", "使用人", "卡位使用者"],
    "卡位月度单价": ["卡位月度单价", "卡位单价", "卡位月单价"],

    # 实验室
    "实验室名称": ["实验室名称", "实验室名", "工作室", "实验室/工作室名称"],
    "实验室编号": ["实验室编号", "实验室号", "LAB编号", "Lab编号"],
    "实验室面积": ["实验室面积", "面积(实验室)"],
    "实验室申请比例": ["实验室申请比例", "申请比例", "占用比例"],
    "实验室公摊率": ["实验室公摊率", "公摊率(实验室)", "实验室公摊"],
    "实验室月度电费": ["实验室电费（按上季度费用核算）", "实验室月度电费（按上季度费用核算）", "实验室月度电费", "实验室电费"],
    "实验室月度单价": ["实验室月度单价", "实验室单价", "实验室月单价"],

    # 负80
    "负80度冰箱名称": ["负80度冰箱名称", "-80度冰箱名称", "冰箱名称"],
    "负80度冰箱编号": ["负80度冰箱编号", "负80编号", "-80编号", "冰箱编号"],
    "负80度冰箱面积": ["负80度冰箱面积", "负80面积", "-80面积", "冰箱面积"],
    "负80度冰箱功率": ["负80度冰箱功率（W）", "负80度冰箱功率", "负80功率", "-80功率", "冰箱功率", "功率（W）", "功率"],
    "负80度每日电量": ["负80度每日电量（度）", "负80度每日电量", "负80每日电量", "-80每日电量", "每日电量（度）", "每日电量"],
    "负80度每月电费": ["负80度每月电费(电费单价1元)", "负80度每月电费", "每月电费(电费单价1元)", "每月电费", 
                      "负80度月度电费(电费单价1元)", "负80度月度电费", "月度电费(电费单价1元)", "月度电费"],
    "负80度冰箱月度单价": ["负80度冰箱月度单价", "负80度月度单价", "负80月度单价", "-80月度单价", "冰箱月度单价"],
}


def _norm_col(x: str) -> str:
    s = "" if x is None else str(x)
    s = s.replace("\u3000", " ").replace("\xa0", " ")
    s = " ".join(s.split()).strip()
    return s


def build_col_map(columns: List[str]) -> Dict[str, str]:
    """返回 {标准名: 实际列名}"""
    norm_to_real = {_norm_col(c): c for c in columns}
    norm_no_space_to_real = {_norm_col(c).replace(" ", ""): c for c in columns}

    out = {}
    for std, alist in ALIASES.items():
        found = None
        for a in alist:
            a1 = _norm_col(a)
            a2 = a1.replace(" ", "")
            if a1 in norm_to_real:
                found = norm_to_real[a1]
                break
            if a2 in norm_no_space_to_real:
                found = norm_no_space_to_real[a2]
                break
            # 宽松：前缀包含
            for k, real in norm_to_real.items():
                if k.startswith(a1) or a1.startswith(k):
                    found = real
                    break
            if found:
                break
        if found:
            out[std] = found
    return out


def parse_date(val) -> Optional[date]:
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    s = str(val).strip()
    if not s:
        return None
    s = s.replace(".", "-").replace("/", "-")
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except Exception:
        try:
            return datetime.fromisoformat(s).date()
        except Exception:
            return None


def normalize_action(x: Any) -> str:
    s = "" if x is None else str(x)
    s = s.replace("\u3000", " ")
    s = re.sub(r"\s+", "", s)
    if "腾出" in s or "腾退" in s:
        return "腾退"
    if "新增" in s:
        return "新增"
    return s


def safe_float(x) -> Optional[float]:
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return None
    try:
        s = str(x).strip()
        if not s:
            return None
        if s.endswith("%"):
            return float(s[:-1]) / 100.0
        return float(s)
    except Exception:
        return None


def _is_valid_text(x) -> bool:
    if x is None:
        return False
    if isinstance(x, float) and math.isnan(x):
        return False
    s = str(x).strip()
    return bool(s) and s.lower() != "nan"


def detect_resource_type(row: Dict[str, Any], colmap: Dict[str, str]) -> Tuple[Optional[str], Optional[str]]:
    """
    返回 (资源类型, 编号)
    资源类型：办公室 / 卡位 / 实验室 / 负80
    """
    def get(std):
        c = colmap.get(std)
        return row.get(c) if c else None

    office = get("办公室编号")
    seat = get("卡位编号")
    lab = get("实验室编号")
    freezer = get("负80度冰箱编号")

    hits = []
    if _is_valid_text(office):
        hits.append(("办公室", str(office).strip().upper()))
    if _is_valid_text(seat):
        hits.append(("卡位", str(seat).strip().upper()))
    if _is_valid_text(lab):
        hits.append(("实验室", str(lab).strip().upper()))
    if _is_valid_text(freezer):
        hits.append(("负80", str(freezer).strip().upper()))

    if len(hits) == 1:
        return hits[0]
    return None, None


@dataclass
class Event:
    key: Tuple[str, str, str, str]  # (中心, 课题组, 资源类型, 编号)
    dt: date                        # 事件发生时间点：新增用start，腾退用end
    action: str                     # 新增/腾退
    # “本次变化”的计费基数信息（用于 + / -）
    # 对办公室/卡位/实验室：base_area_month = 面积*(1+公摊)*单价（已含数量）
    # 对负80：base_elec_month = 每月电费
    base_area_month: float = 0.0
    base_elec_month: float = 0.0

    # 用于输出展示/追溯的字段（尽量保留）
    name: Optional[str] = None
    floor: Optional[Any] = None
    addr: Optional[Any] = None
    handler: Optional[Any] = None
    apply_date: Optional[date] = None

    # 原始字段保留（可选）
    raw_area: Optional[float] = None
    raw_share: Optional[float] = None
    raw_price: Optional[float] = None
    raw_qty: Optional[float] = None
    raw_user: Optional[Any] = None


@dataclass
class State:
    # 当前有效的“月度计费基数”
    area_month: float = 0.0        # 面积相关（办公室/卡位/实验室）
    elec_month: float = 0.0        # 电费相关（实验室/负80）
    # 最近一次可用的展示字段
    name: Optional[str] = None
    floor: Optional[Any] = None
    addr: Optional[Any] = None
    handler: Optional[Any] = None
    # 仅用于展示
    raw_area: Optional[float] = None
    raw_share: Optional[float] = None
    raw_price: Optional[float] = None
    raw_qty: Optional[float] = None
    raw_user: Optional[Any] = None
    # 用于备注：记录事件历史和当前使用面积
    current_area: float = 0.0      # 当前使用面积（累加/减少）
    event_history: List[str] = field(default_factory=list)  # 事件历史列表，格式：["2023-04-28 新增20平", "2023-08-26 新增52平"]


def _compute_days_by_segment(start: date, end: date) -> int:
    """
    计算分段天数（闭区间，含首尾两日）
    统一口径：days = (end - start).days + 1
    """
    d = (end - start).days + 1
    return max(d, 0)


def split_segment_by_month(start_date: date, end_date: date, total_fee: float, total_days: int) -> List[Tuple[str, int, float, str]]:
    """
    将跨月段按天数比例分配到各月（闭区间口径）
    返回：[(year_month, days_in_month, fee_in_month, note), ...]
    year_month格式：'YYYY-MM'
    
    注意：total_days 传入的是闭区间天数（已包含首尾两日），
         拆分后每月 days 之和必须等于 total_days。
    """
    if total_days <= 0:
        return []
    
    from datetime import timedelta
    from collections import defaultdict
    
    # 统计每个月的天数（闭区间：包含 start_date 和 end_date）
    month_days: Dict[str, int] = defaultdict(int)
    current_date = start_date
    
    # 遍历每一天，统计每个月的天数（使用 <= 确保包含 end_date）
    while current_date <= end_date:
        year_month = current_date.strftime('%Y-%m')
        month_days[year_month] += 1
        current_date += timedelta(days=1)
    
    # 验证：每月天数之和应等于 total_days
    sum_days = sum(month_days.values())
    if sum_days != total_days:
        # 如果不一致，使用实际统计的天数（这不应该发生，但作为保护）
        actual_total_days = sum_days
    else:
        actual_total_days = total_days
    
    # 按比例分配费用
    result = []
    months = sorted(month_days.keys())
    
    for year_month in months:
        days_in_month = month_days[year_month]
        fee_in_month = (days_in_month / actual_total_days) * total_fee
        
        # 生成备注
        if len(months) == 1:
            # 不跨月
            note = f"{year_month[-2:]}月{days_in_month}天"
        else:
            # 跨月
            month_name = year_month[-2:] + "月"
            note = f"{month_name}{days_in_month}天(费用{round(fee_in_month, 2)})"
        
        result.append((year_month, days_in_month, fee_in_month, note))
    
    return result


def compute_days(start: Optional[date], end: Optional[date], action: str, today: date) -> Tuple[Optional[int], str]:
    """
    返回 (days, reason)
    统一使用闭区间公式：days = (end - start).days + 1
    - 新增且 end 为空：days = (today - start).days + 1
    - 已腾退：days = (end - start).days + 1
    """
    if not start:
        return None, "缺少计费开始时间"
    if action == "新增" and not end:
        d = (today - start).days + 1
        return max(d, 0), "新增未腾退(截至今天)"
    if not end:
        return None, "缺少计费截止时间"
    d = (end - start).days + 1
    return max(d, 0), "已腾退(按截止-开始)"


def _get(row: Dict[str, Any], colmap: Dict[str, str], std: str):
    c = colmap.get(std)
    return row.get(c) if c else None


def _calc_base_for_row(rtype: str, row: Dict[str, Any], colmap: Dict[str, str]) -> Tuple[float, float, Dict[str, Any], List[str]]:
    """
    返回：
    - base_area_month：按月计费的“面积相关基数”
    - base_elec_month：按月计费的“电费相关基数”
    - extra：用于展示的原始字段（面积/公摊/单价/数量/使用人）
    - issues：本行字段问题
    """
    issues = []
    extra: Dict[str, Any] = {}
    base_area_month = 0.0
    base_elec_month = 0.0

    if rtype == "办公室":
        area = safe_float(_get(row, colmap, "办公室面积"))
        share = safe_float(_get(row, colmap, "办公室公摊率"))
        price = safe_float(_get(row, colmap, "办公室月度单价"))
        extra.update(raw_area=area, raw_share=share, raw_price=price)
        if area is None or share is None or price is None:
            issues.append("办公室字段不完整(面积/公摊率/单价)")
        else:
            base_area_month = area * (1.0 + share) * price

    elif rtype == "卡位":
        qty = safe_float(_get(row, colmap, "卡位数量")) or 1.0
        area = safe_float(_get(row, colmap, "卡位面积"))
        share = safe_float(_get(row, colmap, "卡位公摊率"))
        price = safe_float(_get(row, colmap, "卡位月度单价"))
        user = _get(row, colmap, "卡位使用人")
        extra.update(raw_qty=qty, raw_area=area, raw_share=share, raw_price=price, raw_user=user)
        if area is None or share is None or price is None:
            issues.append("卡位字段不完整(面积/公摊率/单价)")
        else:
            base_area_month = qty * area * (1.0 + share) * price

    elif rtype == "实验室":
        area = safe_float(_get(row, colmap, "实验室面积"))
        share = safe_float(_get(row, colmap, "实验室公摊率"))
        price = safe_float(_get(row, colmap, "实验室月度单价"))
        elec = safe_float(_get(row, colmap, "实验室月度电费"))
        extra.update(raw_area=area, raw_share=share, raw_price=price)
        # 注意：实验室电费是“月度电费”，它本身不需要乘面积；是否按比例分摊由你填表控制
        if area is None or share is None or price is None:
            issues.append("实验室字段不完整(面积/公摊率/单价)")
        else:
            base_area_month = area * (1.0 + share) * price
        if elec is not None:
            base_elec_month = elec

    elif rtype == "负80":
        elec_month = safe_float(_get(row, colmap, "负80度每月电费"))
        if elec_month is None:
            issues.append("负80缺少每月电费")
        else:
            base_elec_month = elec_month

    else:
        issues.append("未知资源类型")

    return float(base_area_month or 0.0), float(base_elec_month or 0.0), extra, issues


def _pick_name_for_row(rtype: str, row: Dict[str, Any], colmap: Dict[str, str]) -> Optional[str]:
    if rtype == "办公室":
        v = _get(row, colmap, "办公室名称")
        return None if not _is_valid_text(v) else str(v).strip()
    if rtype == "卡位":
        v = _get(row, colmap, "卡位名称")
        return None if not _is_valid_text(v) else str(v).strip()
    if rtype == "实验室":
        v = _get(row, colmap, "实验室名称")
        return None if not _is_valid_text(v) else str(v).strip()
    if rtype == "负80":
        v = _get(row, colmap, "负80度冰箱名称")
        return None if not _is_valid_text(v) else str(v).strip()
    return None


def _event_datetime(action: str, start: Optional[date], end: Optional[date]) -> Optional[date]:
    """
    事件发生时间点：
    - 新增：start（计费开始时间）
    - 腾退：end（计费截止时间）
    """
    if action == "新增":
        return start
    if action == "腾退":
        return end
    return None


@dataclass
class ResourceEvent:
    """通用资源事件类"""
    rtype: str  # 资源类型
    center: Optional[str]
    team: str
    code: str
    name: Optional[str]
    addr: Optional[str]
    floor: Optional[Any]
    dt: date
    action: str
    # 资源特定参数（根据资源类型不同而不同）
    area: Optional[float] = None
    share: Optional[float] = None
    price: Optional[float] = None
    qty: Optional[float] = None
    user: Optional[str] = None
    elec: Optional[float] = None
    # 其他字段
    handler: Optional[str] = None
    apply_date: Optional[date] = None
    row_idx: int = 0


def process_resource_events_by_alignment(
    rtype: str,
    rows: List[Dict[str, Any]],
    colmap: Dict[str, str],
    today: date,
    detail_records: List[Dict[str, Any]]
) -> None:
    """
    通用资源事件对齐和分段计费函数
    
    参数:
    - rtype: 资源类型（"办公室"/"卡位"/"实验室"/"负80"）
    - rows: 资源行数据列表，每个元素为 {"row": dict, "idx": int, "is_error": bool, "error_msg": str}
    - colmap: 列名映射
    - today: 今天日期
    - detail_records: 输出明细记录列表（会追加）
    """
    print(f"  [DEBUG] {rtype}行数: {len(rows)}，开始事件对齐分段计费")
    
    events: List[ResourceEvent] = []
    
    # 第一步：从行数据提取事件
    for item in rows:
        row = item["row"]
        idx = item["idx"]
        is_error = item.get("is_error", False)
        error_msg = item.get("error_msg", "")
        
        if is_error:
            # 生成异常记录
            rec = dict(
                资源类型=rtype,
                编号=None,
                所属中心=_get(row, colmap, "所属中心"),
                所属课题组=_get(row, colmap, "所属课题组"),
                名称=None,
                地址=_get(row, colmap, "地址"),
                楼层=_get(row, colmap, "楼层"),
                计费开始时间=parse_date(_get(row, colmap, "计费开始时间")),
                计费截止时间=parse_date(_get(row, colmap, "计费截止时间")),
                使用天数=None,
                天数口径=None,
                经办人=_get(row, colmap, "经办人"),
                费用=None,
                费用计算公式=None,
                异常=error_msg,
                备注=error_msg,
            )
            detail_records.append(rec)
            continue
        
        # 提取通用字段
        center = _get(row, colmap, "所属中心")
        team = _get(row, colmap, "所属课题组")
        handler = _get(row, colmap, "经办人")
        floor = _get(row, colmap, "楼层")
        addr = _get(row, colmap, "地址")
        apply_dt = parse_date(_get(row, colmap, "申请日期"))
        action = normalize_action(_get(row, colmap, "新增/腾退"))
        start = parse_date(_get(row, colmap, "计费开始时间"))
        end = parse_date(_get(row, colmap, "计费截止时间"))
        
        # 提取资源特定字段
        code = None
        name = None
        area = None
        share = None
        price = None
        qty = None
        user = None
        elec = None
        
        if rtype == "办公室":
            code = _get(row, colmap, "办公室编号")
            name = _get(row, colmap, "办公室名称")
            area = safe_float(_get(row, colmap, "办公室面积"))
            share = safe_float(_get(row, colmap, "办公室公摊率"))
            price = safe_float(_get(row, colmap, "办公室月度单价"))
        elif rtype == "卡位":
            code = _get(row, colmap, "卡位编号")
            name = _get(row, colmap, "卡位名称")
            area = safe_float(_get(row, colmap, "卡位面积"))
            share = safe_float(_get(row, colmap, "卡位公摊率"))
            price = safe_float(_get(row, colmap, "卡位月度单价"))
            qty = safe_float(_get(row, colmap, "卡位数量")) or 1.0
            user = _get(row, colmap, "卡位使用人")
        elif rtype == "实验室":
            code = _get(row, colmap, "实验室编号")
            name = _get(row, colmap, "实验室名称")
            area = safe_float(_get(row, colmap, "实验室面积"))
            share = safe_float(_get(row, colmap, "实验室公摊率"))
            price = safe_float(_get(row, colmap, "实验室月度单价"))
            elec = safe_float(_get(row, colmap, "实验室月度电费"))
        elif rtype == "负80":
            code = _get(row, colmap, "负80度冰箱编号")
            name = _get(row, colmap, "负80度冰箱名称")
            elec = safe_float(_get(row, colmap, "负80度每月电费"))
            price = safe_float(_get(row, colmap, "负80度冰箱月度单价"))
            # 负80度冰箱可能还有面积、功率等字段，但目前主要用月度电费和月度单价
        
        # 检查错误情况
        error_msg = None
        if not _is_valid_text(team) or not _is_valid_text(code):
            error_msg = f"行{idx}: 缺少所属课题组或{rtype}编号"
        elif action not in ("新增", "腾退"):
            error_msg = f"行{idx}: 未知的新增/腾退类型: {action}"
        elif action == "新增" and not start:
            error_msg = f"行{idx}: 新增事件缺少计费开始时间"
        elif action == "腾退" and not end:
            error_msg = f"行{idx}: 腾退事件缺少计费截止时间"
        
        if error_msg:
            # 生成异常记录
            rec = dict(
                资源类型=rtype,
                编号=str(code).strip().upper() if code else None,
                所属中心=str(center).strip() if center else None,
                所属课题组=str(team).strip() if team else None,
                名称=str(name).strip() if name else None,
                地址=str(addr).strip() if addr else None,
                楼层=floor,
                计费开始时间=start,
                计费截止时间=end,
                使用天数=None,
                天数口径=None,
                经办人=str(handler).strip() if handler else None,
                面积=area,
                公摊率=share,
                月度单价=price,
                数量=qty,
                使用人=user,
                月度电费=elec if elec else None,
                费用=None,
                费用计算公式=None,
                异常=error_msg,
                备注=error_msg,
            )
            detail_records.append(rec)
            continue
        
        # 确定事件时间
        event_dt = _event_datetime(action, start, end)
        if not event_dt:
            continue
        
        ev = ResourceEvent(
            rtype=rtype,
            center=str(center).strip() if center else None,
            team=str(team).strip(),
            code=str(code).strip().upper(),
            name=str(name).strip() if name else None,
            addr=str(addr).strip() if addr else None,
            floor=floor,
            dt=event_dt,
            action=action,
            area=area,
            share=share,
            price=price,
            qty=qty,
            user=user,
            elec=elec,
            handler=str(handler).strip() if handler else None,
            apply_date=apply_dt,
            row_idx=idx,
        )
        events.append(ev)
        print(f"  [DEBUG] 行{idx}: {rtype}事件已添加 -> {action}, 时间={event_dt}")
    
    if not events:
        return
    
    # 第二步：按资源类型定义的分组key进行分组
    from collections import defaultdict
    
    # 定义每种资源类型的分组key
    def get_group_key(ev: ResourceEvent) -> Tuple:
        if ev.rtype == "办公室":
            # 办公室：所属中心 + 所属课题组 + 办公室名称 + 办公室编号 + 办公室面积 + 办公室公摊率 + 办公室月度单价
            return (ev.center, ev.team, ev.name, ev.code, ev.area, ev.share, ev.price)
        elif ev.rtype == "卡位":
            # 卡位：所属中心 + 所属课题组 + 卡位名称 + 卡位编号 + 卡位面积 + 卡位公摊率 + 卡位使用人 + 卡位月度单价
            return (ev.center, ev.team, ev.name, ev.code, ev.area, ev.share, ev.user, ev.price)
        elif ev.rtype == "实验室":
            # 实验室：所属中心 + 所属课题组 + 实验室编号 + 实验室名称 + 地址 + 楼层
            return (ev.center, ev.team, ev.code, ev.name, ev.addr, ev.floor)
        elif ev.rtype == "负80":
            # 负80度冰箱：所属中心 + 所属课题组 + 负80度冰箱名称 + 负80度冰箱编号 + 负80度每月电费
            return (ev.center, ev.team, ev.name, ev.code, ev.elec)
        else:
            return (ev.center, ev.team, ev.code)
    
    groups: Dict[Tuple, List[ResourceEvent]] = defaultdict(list)
    for ev in events:
        key = get_group_key(ev)
        groups[key].append(ev)
    
    print(f"  [DEBUG] {rtype}事件分组: {len(groups)} 个组")
    
    # 第三步：对每个组进行事件对齐和分段计费
    for key, evs in groups.items():
        print(f"  [DEBUG] 处理{rtype}组: 事件数={len(evs)}")
        
        # 排序：按时间，同日事件先腾退后新增
        evs_sorted = sorted(evs, key=lambda x: (x.dt, 0 if x.action == "腾退" else 1))
        
        # 状态变量（根据资源类型不同而不同）
        if rtype in ("办公室", "卡位", "实验室"):
            # 面积相关资源：维护当前使用面积
            current_area = 0.0
            current_qty = 0.0  # 卡位的当前数量
            param_share: Optional[float] = None
            param_price: Optional[float] = None
            param_elec: Optional[float] = None
            param_qty: Optional[float] = None  # 卡位的参数数量（用于分组）
            param_user: Optional[str] = None
        else:  # 负80
            # 负80度冰箱：维护当前使用状态（0或1）
            current_in_use = False
            param_elec: Optional[float] = None
            param_price: Optional[float] = None  # 负80度冰箱可能也有月度单价
        
        param_center: Optional[str] = None
        param_handler: Optional[str] = None
        param_name: Optional[str] = None
        event_history: List[str] = []
        
        for i, ev in enumerate(evs_sorted):
            # 应用事件：更新状态
            if rtype in ("办公室", "卡位", "实验室"):
                if ev.action == "新增":
                    if ev.area is not None and ev.area > 0:
                        current_area += float(ev.area)
                        if rtype == "卡位" and ev.qty is not None:
                            current_qty += float(ev.qty)
                            event_history.append(f"{ev.dt.strftime('%Y-%m-%d')} 新增{ev.qty}个(面积{ev.area}平)")
                        else:
                            event_history.append(f"{ev.dt.strftime('%Y-%m-%d')} 新增{ev.area}平")
                    else:
                        if rtype == "卡位" and ev.qty is not None:
                            current_qty += float(ev.qty)
                            event_history.append(f"{ev.dt.strftime('%Y-%m-%d')} 新增{ev.qty}个(面积未填)")
                        else:
                            event_history.append(f"{ev.dt.strftime('%Y-%m-%d')} 新增(面积未填)")
                else:  # 腾退
                    if ev.area is not None and ev.area > 0:
                        current_area -= float(ev.area)
                        if rtype == "卡位" and ev.qty is not None:
                            current_qty -= float(ev.qty)
                            event_history.append(f"{ev.dt.strftime('%Y-%m-%d')} 腾退{ev.qty}个(面积{ev.area}平)")
                        else:
                            event_history.append(f"{ev.dt.strftime('%Y-%m-%d')} 腾退{ev.area}平")
                    else:
                        current_area = 0.0
                        if rtype == "卡位":
                            current_qty = 0.0
                        event_history.append(f"{ev.dt.strftime('%Y-%m-%d')} 腾退(全退)")
                
                # 负数保护
                if current_area < 0:
                    issues_msg = f"腾退超过占用(当前{current_area:.1f}平)，已钳制到0"
                    current_area = 0.0
                else:
                    issues_msg = ""
                if rtype == "卡位" and current_qty < 0:
                    issues_msg += f"; 数量为负(当前{current_qty:.0f}个)，已钳制到0"
                    current_qty = 0.0
                
                # 更新参数
                if ev.share is not None:
                    param_share = ev.share
                if ev.price is not None:
                    param_price = ev.price
                if ev.elec is not None:
                    param_elec = ev.elec
                if ev.qty is not None:
                    param_qty = ev.qty  # 用于分组
                if ev.user:
                    param_user = ev.user
            else:  # 负80
                if ev.action == "新增":
                    current_in_use = True
                    event_history.append(f"{ev.dt.strftime('%Y-%m-%d')} 新增")
                else:  # 腾退
                    current_in_use = False
                    event_history.append(f"{ev.dt.strftime('%Y-%m-%d')} 腾退")
                issues_msg = ""
                
                if ev.elec is not None:
                    param_elec = ev.elec
                if ev.price is not None:
                    param_price = ev.price
            
            # 更新通用参数
            if ev.center:
                param_center = ev.center
            if ev.handler:
                param_handler = ev.handler
            if ev.name:
                param_name = ev.name
            
            # 生成分段：段结束日语义（闭区间口径，避免跨段重叠）
            # Retirement is the last billed day of the previous state.
            seg_start = ev.dt + timedelta(days=1) if ev.action == "腾退" else ev.dt
            if i + 1 < len(evs_sorted):
                next_ev = evs_sorted[i + 1]
                # 如果下一事件是腾退，则当前段覆盖到腾退日（业务上"计费截止时间"通常就是最后计费日）
                if next_ev.action == "腾退":
                    seg_end = next_ev.dt
                else:
                    # 否则当前段覆盖到下一事件发生前一天（避免重叠）
                    seg_end = next_ev.dt - timedelta(days=1)
            else:
                seg_end = today
            
            # 检查同日事件（seg_end < seg_start 时，该段天数应为 0）
            if seg_end < seg_start:
                # 跳过该段，不生成记录
                continue
            
            # 生成分段记录
            days = _compute_days_by_segment(seg_start, seg_end)
            print(f"  [DEBUG] 行{ev.row_idx}: 生成段 {seg_start} -> {seg_end}, 天数={days}")
            
            issues = []
            fee = None
            fee_formula = None
            
            # 检查天数是否有效
            if days <= 0:
                if seg_start > today:
                    issues.append(f"计费开始时间为未来日期({seg_start})，无法计费")
                elif seg_end < seg_start:
                    issues.append(f"计费截止时间({seg_end})早于开始时间({seg_start})")
                else:
                    issues.append(f"计费天数无效（{days}天），开始时间={seg_start}，结束时间={seg_end}")
            
            # 计算费用（根据资源类型）
            if rtype == "办公室":
                if param_share is None or param_price is None:
                    issues.append("缺少公摊率或月度单价")
                elif days > 0:
                    if current_area > 0:
                        base_area_month = current_area * (1.0 + param_share) * param_price
                        fee = (base_area_month / DAY_DIVISOR) * days
                        fee_formula = f"费用 = (当前面积 × (1 + 公摊率) × 月度单价 / {DAY_DIVISOR}) × 天数 = ({current_area:.2f} × (1 + {param_share:.4f}) × {param_price:.2f} / {DAY_DIVISOR}) × {days} = {round(fee, 2)}"
                    else:
                        fee = 0.0
                        fee_formula = "当前面积为0，费用 = 0"
            elif rtype == "卡位":
                if param_share is None or param_price is None:
                    issues.append("缺少公摊率或月度单价")
                elif days > 0:
                    if current_area > 0 and current_qty > 0:
                        base_area_month = current_qty * current_area * (1.0 + param_share) * param_price
                        fee = (base_area_month / DAY_DIVISOR) * days
                        fee_formula = f"费用 = (当前数量 × 当前面积 × (1 + 公摊率) × 月度单价 / {DAY_DIVISOR}) × 天数 = ({current_qty:.0f} × {current_area:.2f} × (1 + {param_share:.4f}) × {param_price:.2f} / {DAY_DIVISOR}) × {days} = {round(fee, 2)}"
                    else:
                        fee = 0.0
                        fee_formula = "当前面积或数量为0，费用 = 0"
            elif rtype == "实验室":
                if param_share is None or param_price is None:
                    issues.append("缺少公摊率或月度单价")
                elif days > 0:
                    if current_area > 0:
                        base_area_month = current_area * (1.0 + param_share) * param_price
                        fee = (base_area_month / DAY_DIVISOR) * days
                        fee_formula = f"面积费用 = (当前面积 × (1 + 公摊率) × 月度单价 / {DAY_DIVISOR}) × 天数 = ({current_area:.2f} × (1 + {param_share:.4f}) × {param_price:.2f} / {DAY_DIVISOR}) × {days} = {round(fee, 2)}"
                        
                        if param_elec is not None and param_elec > 0:
                            elec_fee = (param_elec / DAY_DIVISOR) * days
                            fee += elec_fee
                            fee_formula += f" + 电费 = (月度电费 / {DAY_DIVISOR}) × 天数 = ({param_elec:.2f} / {DAY_DIVISOR}) × {days} = {round(elec_fee, 2)}，合计 = {round(fee, 2)}"
                    else:
                        fee = 0.0
                        fee_formula = "当前面积为0，费用 = 0"
            elif rtype == "负80":
                if param_elec is None:
                    issues.append("缺少月度电费")
                elif days > 0:
                    if current_in_use:
                        fee = (param_elec / DAY_DIVISOR) * days
                        fee_formula = f"费用 = (月度电费 / {DAY_DIVISOR}) × 天数 = ({param_elec:.2f} / {DAY_DIVISOR}) × {days} = {round(fee, 2)}"
                    else:
                        fee = 0.0
                        fee_formula = "未使用，费用 = 0"
            
            if issues_msg:
                issues.append(issues_msg)
            
            # 问题1：腾退后的0费用段不应输出
            # 检查是否是最后一个事件，且事件应用后状态为未占用，且seg_end == today
            is_last_event = (i + 1 >= len(evs_sorted))
            is_unoccupied = False
            if rtype in ("办公室", "卡位", "实验室"):
                is_unoccupied = (current_area <= 0)
            elif rtype == "负80":
                is_unoccupied = (not current_in_use)
            
            if is_last_event and is_unoccupied and seg_end == today:
                # 跳过该段，不输出
                print(f"  [DEBUG] 行{ev.row_idx}: 跳过腾退后的0费用段（最后一个事件，未占用状态，截止到today）")
                continue
            
            # 生成备注
            remark = ""
            if event_history:
                event_str = " | ".join(event_history)
                remark = f"事件: {event_str}"
            if rtype in ("办公室", "卡位", "实验室") and current_area > 0:
                if remark:
                    if rtype == "卡位" and current_qty > 0:
                        remark += f" | 当前使用: {current_qty:.0f}个, {current_area:.1f}平"
                    else:
                        remark += f" | 当前使用: {current_area:.1f}平"
                else:
                    if rtype == "卡位" and current_qty > 0:
                        remark = f"当前使用: {current_qty:.0f}个, {current_area:.1f}平"
                    else:
                        remark = f"当前使用: {current_area:.1f}平"
            elif rtype == "负80" and current_in_use:
                if remark:
                    remark += " | 使用中"
                else:
                    remark = "使用中"
            
            # 构建记录
            rec = dict(
                资源类型=rtype,
                编号=ev.code,
                所属中心=param_center,
                所属课题组=ev.team,
                名称=param_name,
                地址=ev.addr,  # 问题2：所有资源类型都输出地址
                楼层=ev.floor,  # 问题2：所有资源类型都输出楼层
                申请日期=ev.apply_date,  # 段起始事件的申请日期
                计费开始时间=seg_start,
                计费截止时间=seg_end,
                新增_腾退=ev.action,  # 段起始事件（触发这段开始的事件：新增/腾退）
                段起始事件时间=ev.dt,  # 段起始事件的发生时间
                使用天数=days,
                天数口径="事件对齐分段" if seg_end < today else "事件对齐分段(截至今天)",
                经办人=param_handler,
                面积=current_area if rtype in ("办公室", "卡位", "实验室") else None,
                公摊率=param_share if rtype in ("办公室", "卡位", "实验室") else None,
                月度单价=param_price if rtype in ("办公室", "卡位", "实验室", "负80") else None,
                数量=current_qty if rtype == "卡位" else None,
                使用人=param_user if rtype == "卡位" else None,
                月度电费=param_elec if param_elec else None,
                费用=round(float(fee), 2) if fee is not None else None,
                费用计算公式=fee_formula,
                异常="; ".join(issues) if issues else "",
                备注=remark if remark else None,
            )
            detail_records.append(rec)
            print(f"  [DEBUG] 行{ev.row_idx}: 已生成段记录，费用={rec.get('费用')}, 异常={rec.get('异常') if rec.get('异常') else '无'}")


# ==================== Stage 1：事件对齐 + 分段计费明细 ====================
def stage1_compute_billing_detail(app_path: str) -> pd.DataFrame:
    """
    实验室：使用状态机分段计费
    办公室/卡位/负80：保持原逻辑，每行独立计费
    """
    df = pd.read_excel(app_path, sheet_name=0, dtype=object)
    df.columns = [_norm_col(c) for c in df.columns]
    colmap = build_col_map(list(df.columns))

    print(f"  [DEBUG] 读取到 {len(df)} 行数据")
    print(f"  [DEBUG] 关键列映射：")
    for k in ["所属中心", "所属课题组", "新增/腾退", "计费开始时间", "计费截止时间",
              "办公室编号", "卡位编号", "实验室编号", "负80度冰箱编号"]:
        print(f"    {k} -> {colmap.get(k, '未找到')}")

    today = date.today()
    
    # 按资源类型分类
    office_rows: List[Dict[str, Any]] = []
    seat_rows: List[Dict[str, Any]] = []
    lab_rows: List[Dict[str, Any]] = []
    freezer_rows: List[Dict[str, Any]] = []
    error_rows: List[Dict[str, Any]] = []
    
    print(f"  [DEBUG] 开始处理 {len(df)} 行数据")
    for idx, r in df.iterrows():
        row = r.to_dict()
        rtype, code = detect_resource_type(row, colmap)
        excel_row = idx + 2  # Excel行号（从2开始，因为第1行是表头）
        if not rtype:
            # 无法识别资源类型，生成异常记录
            print(f"  [DEBUG] 行{excel_row}: 无法识别资源类型 -> 将生成异常记录")
            error_rows.append({
                "row": row, 
                "idx": excel_row,
                "is_error": True,
                "error_msg": "无法识别资源类型（所有编号字段为空或冲突）"
            })
            continue
        
        print(f"  [DEBUG] 行{excel_row}: 识别为 {rtype}, 编号={code}")
        if rtype == "办公室":
            office_rows.append({"row": row, "idx": excel_row, "is_error": False})
        elif rtype == "卡位":
            seat_rows.append({"row": row, "idx": excel_row, "is_error": False})
        elif rtype == "实验室":
            lab_rows.append({"row": row, "idx": excel_row, "is_error": False})
        elif rtype == "负80":
            freezer_rows.append({"row": row, "idx": excel_row, "is_error": False})
        else:
            error_rows.append({
                "row": row,
                "idx": excel_row,
                "is_error": True,
                "error_msg": f"未知资源类型: {rtype}"
            })
    
    print(f"  [DEBUG] 分类结果: 办公室={len(office_rows)}行, 卡位={len(seat_rows)}行, 实验室={len(lab_rows)}行, 负80={len(freezer_rows)}行, 错误={len(error_rows)}行")
    
    detail_records: List[Dict[str, Any]] = []
    
    # ========== 处理错误行（无法识别资源类型） ==========
    print(f"  [DEBUG] 开始处理 {len(error_rows)} 行错误行")
    for item in error_rows:
        row = item["row"]
        idx = item["idx"]
        error_msg = item.get("error_msg", "")
        print(f"  [DEBUG] 行{idx}: 生成异常记录 -> {error_msg}")
        rec = dict(
            资源类型=None,
            编号=None,
            所属中心=_get(row, colmap, "所属中心"),
            所属课题组=_get(row, colmap, "所属课题组"),
            名称=None,
            地址=_get(row, colmap, "地址"),
            楼层=_get(row, colmap, "楼层"),
            申请日期=parse_date(_get(row, colmap, "申请日期")),
            计费开始时间=parse_date(_get(row, colmap, "计费开始时间")),
            计费截止时间=parse_date(_get(row, colmap, "计费截止时间")),
            新增_腾退=normalize_action(_get(row, colmap, "新增/腾退")),
            使用天数=None,
            天数口径=None,
            经办人=_get(row, colmap, "经办人"),
            面积=None,
            公摊率=None,
            月度单价=None,
            数量=None,
            使用人=None,
            月度电费=None,
            费用=None,
            费用计算公式=None,
            异常=error_msg,
            备注=f"行{idx}: {error_msg}",
        )
        detail_records.append(rec)
    
    # ========== 使用事件对齐处理所有资源类型 ==========
    # 办公室
    if office_rows:
        process_resource_events_by_alignment("办公室", office_rows, colmap, today, detail_records)
    
    # 卡位
    if seat_rows:
        process_resource_events_by_alignment("卡位", seat_rows, colmap, today, detail_records)
    
    # 实验室
    if lab_rows:
        process_resource_events_by_alignment("实验室", lab_rows, colmap, today, detail_records)
    
    # 负80度冰箱
    if freezer_rows:
        process_resource_events_by_alignment("负80", freezer_rows, colmap, today, detail_records)
    
    # 以下为旧的实验室处理代码（已废弃，保留作为参考）
    if False and lab_rows:
        print(f"  [DEBUG] 实验室行数: {len(lab_rows)}，开始状态机分段计费")
        
        # 定义实验室事件key：所属课题组 + 实验室编号(优先) + 实验室名称(辅助) + 地址 + 楼层(可选)
        @dataclass
        class LabEvent:
            team: str
            code: str
            name: Optional[str]
            addr: Optional[str]
            floor: Optional[Any]
            dt: date
            action: str
            area: Optional[float]  # 实验室面积
            share: Optional[float]  # 公摊率
            price: Optional[float]  # 月度单价
            elec: Optional[float]  # 月度电费
            center: Optional[str]
            handler: Optional[str]
            apply_date: Optional[date]
            row_idx: int
        
        lab_events: List[LabEvent] = []
        
        for item in lab_rows:
            row = item["row"]
            idx = item["idx"]
            
            team = _get(row, colmap, "所属课题组")
            code = _get(row, colmap, "实验室编号")
            name = _get(row, colmap, "实验室名称")
            addr = _get(row, colmap, "地址")
            floor = _get(row, colmap, "楼层")
            center = _get(row, colmap, "所属中心")
            handler = _get(row, colmap, "经办人")
            apply_dt = parse_date(_get(row, colmap, "申请日期"))
            action = normalize_action(_get(row, colmap, "新增/腾退"))
            start = parse_date(_get(row, colmap, "计费开始时间"))
            end = parse_date(_get(row, colmap, "计费截止时间"))
            area = safe_float(_get(row, colmap, "实验室面积"))
            share = safe_float(_get(row, colmap, "实验室公摊率"))
            price = safe_float(_get(row, colmap, "实验室月度单价"))
            elec = safe_float(_get(row, colmap, "实验室月度电费"))
            
            # 检查各种错误情况，生成异常记录而不是跳过
            error_msg = None
            if not _is_valid_text(team) or not _is_valid_text(code):
                error_msg = f"行{idx}: 缺少所属课题组或实验室编号"
            elif action not in ("新增", "腾退"):
                error_msg = f"行{idx}: 未知的新增/腾退类型: {action}"
            elif action == "新增" and not start:
                error_msg = f"行{idx}: 新增事件缺少计费开始时间"
            elif action == "腾退" and not end:
                error_msg = f"行{idx}: 腾退事件缺少计费截止时间"
            
            if error_msg:
                # 生成异常记录
                print(f"  [DEBUG] 行{idx}: 实验室异常 -> {error_msg}")
                rec = dict(
                    资源类型="实验室",
                    编号=str(code).strip().upper() if code else None,
                    所属中心=str(center).strip() if center else None,
                    所属课题组=str(team).strip() if team else None,
                    名称=str(name).strip() if name else None,
                    地址=str(addr).strip() if addr else None,
                    楼层=floor,
                    计费开始时间=start,
                    计费截止时间=end,
                    使用天数=None,
                    天数口径=None,
                    经办人=str(handler).strip() if handler else None,
                    面积=area,
                    公摊率=share,
                    月度单价=price,
                    月度电费=elec if elec else None,
                    费用=None,
                    费用计算公式=None,
                    异常=error_msg,
                    备注=error_msg,
                )
                detail_records.append(rec)
                continue
            
            # 正常处理：确定事件时间
            if action == "新增":
                event_dt = start
            else:  # 腾退
                event_dt = end
            
            ev = LabEvent(
                team=str(team).strip(),
                code=str(code).strip().upper(),
                name=str(name).strip() if name else None,
                addr=str(addr).strip() if addr else None,
                floor=floor,
                dt=event_dt,
                action=action,
                area=area,
                share=share,
                price=price,
                elec=elec,
                center=str(center).strip() if center else None,
                handler=str(handler).strip() if handler else None,
                apply_date=apply_dt,
                row_idx=idx,
            )
            lab_events.append(ev)
            print(f"  [DEBUG] 行{idx}: 实验室事件已添加 -> {action}, 时间={event_dt}, 面积={area}")
        
        # 按key分组：所属中心 + 所属课题组 + 实验室编号 + 实验室名称 + 地址 + 楼层
        # 注意：必须包含所属中心，因为注释要求：同一 (所属中心, 所属课题组, 资源类型, 编号) 形成一条时间线
        from collections import defaultdict
        lab_groups: Dict[Tuple[Optional[str], str, str, Optional[str], Optional[str], Optional[Any]], List[LabEvent]] = defaultdict(list)
        for ev in lab_events:
            key = (ev.center, ev.team, ev.code, ev.name, ev.addr, ev.floor)
            lab_groups[key].append(ev)
        
        # 对每个key，按时间排序并分段计费
        print(f"  [DEBUG] 实验室事件分组: {len(lab_groups)} 个组")
        for key, evs in lab_groups.items():
            center, team, code, name, addr, floor = key
            print(f"  [DEBUG] 处理实验室组: 中心={center}, 课题组={team}, 编号={code}, 事件数={len(evs)}")
            
            # 排序：按时间，同日事件先腾退后新增
            evs_sorted = sorted(evs, key=lambda x: (x.dt, 0 if x.action == "腾退" else 1))
            
            # 状态变量
            current_area = 0.0
            param_share: Optional[float] = None
            param_price: Optional[float] = None
            param_elec: Optional[float] = None
            param_center: Optional[str] = None
            param_handler: Optional[str] = None
            param_name: Optional[str] = None
            event_history: List[str] = []
            
            for i, ev in enumerate(evs_sorted):
                # 应用事件：更新current_area
                if ev.action == "新增":
                    if ev.area is not None and ev.area > 0:
                        current_area += float(ev.area)
                        event_history.append(f"{ev.dt.strftime('%Y-%m-%d')} 新增{ev.area}平")
                    else:
                        event_history.append(f"{ev.dt.strftime('%Y-%m-%d')} 新增(面积未填)")
                else:  # 腾退
                    if ev.area is not None and ev.area > 0:
                        current_area -= float(ev.area)
                        event_history.append(f"{ev.dt.strftime('%Y-%m-%d')} 腾退{ev.area}平")
                    else:
                        # 全退
                        current_area = 0.0
                        event_history.append(f"{ev.dt.strftime('%Y-%m-%d')} 腾退(全退)")
                
                # 负数保护
                if current_area < 0:
                    issues_msg = f"腾退超过占用(当前{current_area:.1f}平)，已钳制到0"
                    current_area = 0.0
                else:
                    issues_msg = ""
                
                # 更新参数状态（如果事件提供了参数）
                if ev.share is not None:
                    param_share = ev.share
                if ev.price is not None:
                    param_price = ev.price
                if ev.elec is not None:
                    param_elec = ev.elec
                if ev.center:
                    param_center = ev.center
                if ev.handler:
                    param_handler = ev.handler
                if ev.name:
                    param_name = ev.name
                
                # 生成分段：段结束日语义（闭区间口径，避免跨段重叠）
                # Retirement is the last billed day of the previous state.
                seg_start = ev.dt + timedelta(days=1) if ev.action == "腾退" else ev.dt
                if i + 1 < len(evs_sorted):
                    next_ev = evs_sorted[i + 1]
                    # 如果下一事件是腾退，则当前段覆盖到腾退日（业务上"计费截止时间"通常就是最后计费日）
                    if next_ev.action == "腾退":
                        seg_end = next_ev.dt
                    else:
                        # 否则当前段覆盖到下一事件发生前一天（避免重叠）
                        seg_end = next_ev.dt - timedelta(days=1)
                else:
                    seg_end = today
                
                # 检查同日事件（seg_end < seg_start 时，该段天数应为 0）
                if seg_end < seg_start:
                    # 跳过该段，不生成记录
                    continue
                
                seg_area = current_area  # 应用事件后的面积
                
                # 生成分段
                days = _compute_days_by_segment(seg_start, seg_end)
                print(f"  [DEBUG] 行{ev.row_idx}: 生成段 {seg_start} -> {seg_end}, 天数={days}, 面积={seg_area}")
                
                issues = []
                fee = None
                
                # 检查天数是否有效
                if days <= 0:
                    if seg_start > today:
                        issues.append(f"计费开始时间为未来日期({seg_start})，无法计费")
                    elif seg_end < seg_start:
                        issues.append(f"计费截止时间({seg_end})早于开始时间({seg_start})")
                    else:
                        issues.append(f"计费天数无效（{days}天），开始时间={seg_start}，结束时间={seg_end}")
                
                # 检查参数完整性
                fee_formula = None
                if param_share is None or param_price is None:
                    issues.append("缺少公摊率或月度单价")
                elif days > 0:  # 只有天数>0时才计算费用
                    # 费用计算：segment_area * (1 + 公摊率) * 月度单价 / 30 * days
                    if seg_area > 0:
                        base_area_month = seg_area * (1.0 + param_share) * param_price
                        fee = (base_area_month / DAY_DIVISOR) * days
                        fee_formula = f"面积费用 = (当前面积 × (1 + 公摊率) × 月度单价 / {DAY_DIVISOR}) × 天数 = ({seg_area:.2f} × (1 + {param_share:.4f}) × {param_price:.2f} / {DAY_DIVISOR}) × {days} = {round(fee, 2)}"
                        
                        # 如果月度电费非空，加上电费
                        if param_elec is not None and param_elec > 0:
                            elec_fee = (param_elec / DAY_DIVISOR) * days
                            fee += elec_fee
                            fee_formula += f" + 电费 = (月度电费 / {DAY_DIVISOR}) × 天数 = ({param_elec:.2f} / {DAY_DIVISOR}) × {days} = {round(elec_fee, 2)}，合计 = {round(fee, 2)}"
                    else:
                        # 面积为0，费用为0
                        fee = 0.0
                        fee_formula = "当前面积为0，费用 = 0"
                # 如果days<=0，fee保持为None
                
                if issues_msg:
                    issues.append(issues_msg)
                
                # 生成备注
                remark = ""
                if event_history:
                    event_str = " | ".join(event_history)
                    remark = f"事件: {event_str}"
                if seg_area > 0:
                    if remark:
                        remark += f" | 当前使用: {seg_area:.1f}平"
                    else:
                        remark = f"当前使用: {seg_area:.1f}平"
                
                rec = dict(
                    资源类型="实验室",
                    编号=code,
                    所属中心=param_center,
                    所属课题组=team,
                    名称=param_name,
                    地址=addr,
                    楼层=floor,
                    计费开始时间=seg_start,
                    计费截止时间=seg_end,
                    使用天数=days,
                    天数口径="事件对齐分段" if seg_end < today else "事件对齐分段(截至今天)",
                    经办人=param_handler,
                    面积=seg_area,  # segment_area
                    公摊率=param_share,
                    月度单价=param_price,
                    月度电费=param_elec if param_elec else None,
                    费用=round(float(fee), 2) if fee is not None else None,
                    费用计算公式=fee_formula,
                    异常="; ".join(issues) if issues else "",
                    备注=remark if remark else None,
                )
                detail_records.append(rec)
                print(f"  [DEBUG] 行{ev.row_idx}: 已生成段记录，费用={rec.get('费用')}, 异常={rec.get('异常') if rec.get('异常') else '无'}")
    
    # 写入原始问题到异常列（这些记录已经在处理过程中直接添加到detail_records了，这里不再重复添加）
    # 注意：所有跳过的行都已经在处理过程中生成了异常记录，所以这里不需要再添加
    
    print(f"  [DEBUG] 最终明细记录数: {len(detail_records)}")
    detail = pd.DataFrame(detail_records)
    
    # 日期列统一
    for c in ["计费开始时间", "计费截止时间", "申请日期"]:
        if c in detail.columns:
            detail[c] = pd.to_datetime(detail[c], errors="coerce").dt.date
    
    # 费用保留2位
    if "费用" in detail.columns:
        detail["费用"] = pd.to_numeric(detail.get("费用"), errors="coerce").round(2)
    
    return detail


# ==================== 月度统计功能 ====================
def generate_monthly_statistics(detail: pd.DataFrame) -> pd.DataFrame:
    """
    生成月度统计表
    行：所属中心 + 所属课题组 + 年月
    列：办公室、卡位、实验室、负80、合计、备注
    """
    if detail.empty:
        return pd.DataFrame()
    
    # 只处理费用非空的记录
    calc = detail.copy()
    calc["费用"] = pd.to_numeric(calc.get("费用"), errors="coerce")
    calc_valid = calc[calc["费用"].notna() & (calc["费用"] > 0)].copy()
    
    if calc_valid.empty:
        return pd.DataFrame(columns=["所属中心", "所属课题组", "年月", "办公室", "卡位", "实验室", "负80", "合计", "备注"])
    
    monthly_records = []
    
    # 遍历每条明细记录
    for idx, row in calc_valid.iterrows():
        center = row.get("所属中心")
        team = row.get("所属课题组")
        resource_type = row.get("资源类型")
        start_date = row.get("计费开始时间")
        end_date = row.get("计费截止时间")
        fee = row.get("费用")
        days = row.get("使用天数")
        
        if not start_date or not end_date or pd.isna(start_date) or pd.isna(end_date):
            continue
        
        if not isinstance(start_date, date):
            try:
                start_date = pd.to_datetime(start_date).date()
            except:
                continue
        if not isinstance(end_date, date):
            try:
                end_date = pd.to_datetime(end_date).date()
            except:
                continue
        
        if not days or pd.isna(days) or days <= 0:
            continue
        
        # 将段按月份拆分
        month_splits = split_segment_by_month(start_date, end_date, float(fee), int(days))
        
        # 收集该段的所有月份信息，用于生成备注
        if month_splits:
            # 判断是否跨月
            is_cross_month = len(month_splits) > 1
            note_parts = []
            
            for year_month, days_in_month, fee_in_month, note_part in month_splits:
                # 查找是否已有该课题组该月份的记录
                found = False
                for rec in monthly_records:
                    if (rec["所属中心"] == center and 
                        rec["所属课题组"] == team and 
                        rec["年月"] == year_month):
                        # 累加费用
                        rec[resource_type] = rec.get(resource_type, 0.0) + fee_in_month
                        rec["合计"] = rec.get("合计", 0.0) + fee_in_month
                        found = True
                        break
                
                if not found:
                    # 创建新记录
                    new_rec = {
                        "所属中心": center,
                        "所属课题组": team,
                        "年月": year_month,
                        "办公室": 0.0,
                        "卡位": 0.0,
                        "实验室": 0.0,
                        "负80": 0.0,
                        "合计": 0.0,
                        "备注": "",
                    }
                    new_rec[resource_type] = fee_in_month
                    new_rec["合计"] = fee_in_month
                    monthly_records.append(new_rec)
                
                # 收集备注信息
                month_name = year_month[-2:] + "月"
                if is_cross_month:
                    note_parts.append(f"{month_name}{days_in_month}天(费用{round(fee_in_month, 2)})")
                else:
                    note_parts.append(f"{month_name}{days_in_month}天")
            
            # 生成备注
            if is_cross_month:
                full_note = f"跨月：{', '.join(note_parts)}"
            else:
                full_note = note_parts[0] if note_parts else ""
            
            # 为所有涉及的记录更新备注（每个资源类型单独记录）
            for rec in monthly_records:
                if (rec["所属中心"] == center and 
                    rec["所属课题组"] == team and 
                    rec["年月"] in [ym for ym, _, _, _ in month_splits]):
                    if full_note:
                        # 检查是否已存在相同资源类型的备注
                        existing_notes = rec["备注"].split("；") if rec["备注"] else []
                        # 移除相同资源类型的旧备注（通过资源类型前缀识别）
                        resource_prefix = f"{resource_type}:"
                        existing_notes = [n for n in existing_notes if not n.startswith(resource_prefix)]
                        # 添加新备注（格式：资源类型:备注内容）
                        existing_notes.append(f"{resource_prefix}{full_note}")
                        rec["备注"] = "；".join(existing_notes)
    
    if not monthly_records:
        return pd.DataFrame(columns=["所属中心", "所属课题组", "年月", "办公室", "卡位", "实验室", "负80", "合计", "备注"])
    
    # 转换为DataFrame
    monthly_df = pd.DataFrame(monthly_records)
    
    # 按中心、课题组、年月排序
    monthly_df = monthly_df.sort_values(["所属中心", "所属课题组", "年月"])
    
    # 费用保留2位小数
    for col in ["办公室", "卡位", "实验室", "负80", "合计"]:
        if col in monthly_df.columns:
            monthly_df[col] = monthly_df[col].round(2)
    
    return monthly_df


# ==================== Stage 2：生成报表（汇总 + 分类型明细） ====================
def stage2_generate_report(detail: pd.DataFrame, out_path: str) -> None:
    if detail.empty or "资源类型" not in detail.columns:
        print(f"  ✗ 错误：计费明细为空或缺少'资源类型'列，无法生成报表")
        return

    # 只汇总“费用非空”的记录
    calc = detail.copy()
    calc["费用"] = pd.to_numeric(calc.get("费用"), errors="coerce")
    calc_valid = calc[calc["费用"].notna()].copy()

    office = calc_valid[calc_valid["资源类型"] == "办公室"].copy()
    seat = calc_valid[calc_valid["资源类型"] == "卡位"].copy()
    lab = calc_valid[calc_valid["资源类型"] == "实验室"].copy()
    freezer = calc_valid[calc_valid["资源类型"] == "负80"].copy()

    pivot = (
        calc_valid.pivot_table(
            index=["所属中心", "所属课题组"],
            columns=["资源类型"],
            values="费用",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
    )

    for col in ["办公室", "卡位", "实验室", "负80"]:
        if col not in pivot.columns:
            pivot[col] = 0.0
    pivot["合计"] = (pivot["办公室"] + pivot["卡位"] + pivot["实验室"] + pivot["负80"]).round(2)
    for col in ["办公室", "卡位", "实验室", "负80"]:
        pivot[col] = pivot[col].round(2)

    def pick_cols(df, cols):
        exist = [c for c in cols if c in df.columns]
        return df[exist].copy()

    office_cols = [
        "所属中心", "所属课题组",
        "名称", "编号", "楼层", "地址",
        "面积", "公摊率", "月度单价",
        "计费开始时间", "计费截止时间", "使用天数", "费用", "费用计算公式",
        "经办人", "异常",
    ]
    seat_cols = [
        "所属中心", "所属课题组",
        "名称", "编号", "楼层", "地址",
        "数量", "面积", "公摊率", "月度单价",
        "使用人",
        "计费开始时间", "计费截止时间", "使用天数", "费用", "费用计算公式",
        "经办人", "异常",
    ]
    lab_cols = [
        "所属中心", "所属课题组",
        "名称", "编号", "楼层", "地址",
        "面积", "公摊率", "月度电费", "月度单价",
        "计费开始时间", "计费截止时间", "使用天数", "费用", "费用计算公式",
        "经办人", "备注", "异常",
    ]
    freezer_cols = [
        "所属中心", "所属课题组",
        "名称", "编号", "楼层", "地址",
        "月度电费",
        "计费开始时间", "计费截止时间", "使用天数", "费用", "费用计算公式",
        "经办人", "异常",
    ]

    # 生成月度统计表
    monthly_stats = generate_monthly_statistics(detail)
    
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        pivot.to_excel(writer, sheet_name="汇总(中心-课题组)", index=False)
        pick_cols(office, office_cols).to_excel(writer, sheet_name="办公室明细(按段)", index=False)
        pick_cols(seat, seat_cols).to_excel(writer, sheet_name="卡位明细(按段)", index=False)
        pick_cols(lab, lab_cols).to_excel(writer, sheet_name="实验室明细(按段)", index=False)
        pick_cols(freezer, freezer_cols).to_excel(writer, sheet_name="负80明细(按段)", index=False)
        detail.to_excel(writer, sheet_name="计费明细(全量-按段)", index=False)
        if not monthly_stats.empty:
            monthly_stats.to_excel(writer, sheet_name="月度统计(课题组-月份)", index=False)
    
    # 自动调整列宽
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    
    wb = load_workbook(out_path)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        # 获取表头行（第一行）
        if ws.max_row == 0:
            continue
        
        # 遍历每一列
        for col_idx, col in enumerate(ws.iter_cols(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column), start=1):
            max_length = 0
            column_letter = get_column_letter(col_idx)
            
            # 遍历该列的所有单元格
            for cell in col:
                try:
                    # 获取单元格内容长度
                    if cell.value:
                        cell_value = str(cell.value)
                        # 计算显示长度（中文字符按2个字符宽度计算）
                        display_length = 0
                        for char in cell_value:
                            if ord(char) > 127:  # 中文字符
                                display_length += 2
                            else:
                                display_length += 1
                        max_length = max(max_length, display_length)
                except:
                    pass
            
            # 设置列宽（最小宽度10，最大宽度50，中文字符需要额外空间）
            adjusted_width = min(max(max_length + 2, 10), 50)
            ws.column_dimensions[column_letter].width = adjusted_width
    
    wb.save(out_path)
    wb.close()


def main():
    print("=" * 60)
    print("Stage 1：事件对齐 -> 分段计费明细")
    print("=" * 60)
    detail = stage1_compute_billing_detail(APPLICATION_FILE)
    print(f"  ✓ 明细行数：{len(detail)}")
    if "异常" in detail.columns:
        bad = detail[detail["异常"].astype(str).str.strip() != ""]
        print(f"  ⚠ 异常记录数：{len(bad)}（可到输出Excel的“计费明细(全量-按段)”里看）")

    print("\n" + "=" * 60)
    print("Stage 2：生成输出报表（汇总 + 分类型明细）")
    print("=" * 60)
    stage2_generate_report(detail, OUTPUT_REPORT_FILE)
    print(f"  ✓ 已输出：{OUTPUT_REPORT_FILE}")
    print("\n完成。")


if __name__ == "__main__":
    main()
