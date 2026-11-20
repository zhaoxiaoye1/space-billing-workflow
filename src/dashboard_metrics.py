"""Dashboard transformations reused from the delivered administrative billing tool.
Sample data and resource totals are fictional. Resource usage is a count of
distinct identifiers appearing in a monthly billing period, not area utilization.
"""
import json
from datetime import date, datetime
from calendar import monthrange
from typing import Any, Dict, List, Optional, Tuple
import pandas as pd

def get_month_intersection_days(start_date: date, end_date: date, year_month: str, target_month: str = None) -> int:
    """
    计算计费段与指定月份的交集天数
    例如：2025-11月，计算从2025-11-01到2025-11-30与[start_date, end_date]的交集天数
    
    参数:
        target_month: 目标月份（格式：YYYY-MM），如果指定且与year_month相同，则截至今天
    """
    year, month = map(int, year_month.split('-'))
    today = date.today()
    
    # 月份的第一天和最后一天
    month_start = date(year, month, 1)
    month_end = date(year, month, monthrange(year, month)[1])
    
    # 如果目标月份是当前月份，则截至今天
    if target_month == year_month and today.year == year and today.month == month:
        month_end = min(month_end, today)
    
    # 计算交集
    intersection_start = max(start_date, month_start)
    intersection_end = min(end_date, month_end)
    
    if intersection_start > intersection_end:
        return 0
    
    # 计算天数（含首日：需要+1）
    days = (intersection_end - intersection_start).days + 1
    return max(days, 0)

def split_by_month(detail_df: pd.DataFrame, selected_months: List[str], target_month: str = None) -> pd.DataFrame:
    """
    基于 bill2.0 返回的明细数据，按月拆分
    费用按天数比例拆分：当月费用 = 原始费用 × (当月天数 / 总使用天数)
    
    参数:
        target_month: 目标月份（格式：YYYY-MM），用于确定是否截至今天
    """
    monthly_records = []
    today = date.today()
    
    for idx, row in detail_df.iterrows():
        start_date = row.get('计费开始时间')
        end_date = row.get('计费截止时间')
        original_fee = row.get('费用', 0.0)
        original_days = row.get('使用天数', 0)
        
        # 处理费用和天数为 None 或 NaN 的情况
        if pd.isna(original_fee) or original_fee is None:
            original_fee = 0.0
        else:
            try:
                original_fee = float(original_fee)
            except (ValueError, TypeError):
                original_fee = 0.0
        
        if pd.isna(original_days) or original_days is None:
            original_days = 0
        else:
            try:
                original_days = int(original_days)
            except (ValueError, TypeError):
                original_days = 0
        
        if pd.isna(start_date) or original_days <= 0:
            continue
        
        if not isinstance(start_date, date):
            try:
                start_date = pd.to_datetime(start_date).date()
            except:
                continue
        
        if pd.isna(end_date) or end_date is None:
            end_date = today
        else:
            if not isinstance(end_date, date):
                try:
                    end_date = pd.to_datetime(end_date).date()
                except:
                    continue
        
        # 限制 end_date 不超过目标月份的最后一天
        if target_month:
            target_year, target_month_num = map(int, target_month.split('-'))
            target_month_end = date(target_year, target_month_num, monthrange(target_year, target_month_num)[1])
            # 如果目标月份是当前月份，则截至今天
            if target_year == today.year and target_month_num == today.month:
                target_month_end = min(target_month_end, today)
            # 限制 end_date 不超过目标月份
            end_date = min(end_date, target_month_end)
        
        # 为每个选中的月份生成一条记录
        for year_month in selected_months:
            days_in_month = get_month_intersection_days(start_date, end_date, year_month, target_month)
            
            if days_in_month <= 0:
                continue
            
            # 按天数比例拆分费用（基于 bill2.0 返回的原始费用）
            if original_days > 0:
                month_fee = (original_fee / original_days) * days_in_month
            else:
                month_fee = 0.0
            
            # 创建月度记录
            monthly_record = {
                '所区': row.get('地址', ''),
                '所属中心': row.get('所属中心', ''),
                '所属课题组': row.get('所属课题组', ''),
                '资源类型': row.get('资源类型', ''),
                '名称': row.get('名称', ''),
                '编号': row.get('编号', ''),
                '年月': year_month,
                '当月金额': round(month_fee, 2),
                '当月天数': days_in_month,
                '备注': f"{year_month[-2:]}月{days_in_month}天",
                '异常': row.get('异常', ''),
            }
            monthly_records.append(monthly_record)
    
    if not monthly_records:
        return pd.DataFrame()
    
    monthly_df = pd.DataFrame(monthly_records)
    
    # 按资源+月份分组，合并同一资源在同一月份的多条记录
    grouped = monthly_df.groupby(['所属课题组', '所区', '所属中心', '资源类型', '名称', '编号', '年月'], dropna=False)
    
    merged_records = []
    for (team, addr, center, rtype, name, code, year_month), group in grouped:
        total_amount = group['当月金额'].sum()
        total_days = group['当月天数'].sum()
        
        # 合并备注天数信息
        days_info = group['备注'].tolist()
        seen = set()
        unique_days_info = []
        for info in days_info:
            if info not in seen:
                seen.add(info)
                unique_days_info.append(info)
        remark = '，'.join(unique_days_info)
        
        # 合并异常信息
        exceptions = group['异常'].dropna().unique()
        exception_str = '; '.join([str(e) for e in exceptions if str(e).strip()])
        
        merged_record = {
            '所区': addr if pd.notna(addr) else '',
            '所属中心': center if pd.notna(center) else '',
            '所属课题组': team if pd.notna(team) else '',
            '资源类型': rtype if pd.notna(rtype) else '',
            '名称': name if pd.notna(name) else '',
            '编号': code if pd.notna(code) else '',
            '年月': year_month,
            '当月金额': round(total_amount, 2),
            '当月天数': total_days,
            '备注': remark,
            '异常': exception_str,
        }
        merged_records.append(merged_record)
    
    return pd.DataFrame(merged_records)

def calculate_kpis(summary_df: pd.DataFrame, target_month: str, resource_totals: Dict[str, int] = None) -> Dict[str, Any]:
    """
    计算6个核心KPI指标
    
    参数:
        resource_totals: 资源总数配置，用于计算占用率
    """
    if resource_totals is None:
        resource_totals = {'办公室': 100, '卡位': 100, '实验室': 100}
    
    # 筛选目标月份的数据
    month_data = summary_df[summary_df['年月'] == target_month].copy()
    
    if month_data.empty:
        return {
            '本月总费用': 0.0,
            '本月各中心费用': {},
            '本月TOP1课题组费用': {'课题组': '', '金额': 0.0},
            '本月实验室费用占比': 0.0,
            '本月在用资源数': {'办公室': 0, '卡位': 0, '实验室': 0, '负80': 0},
            '本月资源占用率': {'办公室': 0.0, '卡位': 0.0, '实验室': 0.0, '负80': 0.0},
            '本月异常条数': 0,
            '本月异常率': 0.0,
        }
    
    # 1. 本月总费用（全所区）
    total_fee = month_data['当月金额'].sum()
    
    # 2. 本月各中心费用
    center_fees = month_data.groupby('所属中心')['当月金额'].sum().to_dict()
    
    # 3. 本月TOP1课题组费用
    team_fees = month_data.groupby('所属课题组')['当月金额'].sum().sort_values(ascending=False)
    top1_team = team_fees.index[0] if len(team_fees) > 0 else ''
    top1_amount = team_fees.iloc[0] if len(team_fees) > 0 else 0.0
    
    # 4. 本月实验室费用占比
    lab_fee = month_data[month_data['资源类型'] == '实验室']['当月金额'].sum()
    lab_ratio = (lab_fee / total_fee * 100) if total_fee > 0 else 0.0
    
    # 5. 本月在用资源数（按资源类型统计，同一编号只算一次）
    # 按资源类型和编号去重，确保同一编号在同一月份只算一次
    unique_resources = month_data.groupby(['资源类型', '编号']).size().reset_index(name='count')
    resource_counts = unique_resources.groupby('资源类型').size().to_dict()
    resource_stats = {
        '办公室': resource_counts.get('办公室', 0),
        '卡位': resource_counts.get('卡位', 0),
        '实验室': resource_counts.get('实验室', 0),
        '负80': resource_counts.get('负80', 0),
    }
    
    # 计算资源占用率（负80冰箱不计算占比，只显示占用数量）
    resource_usage_rates = {
        '办公室': (resource_stats['办公室'] / resource_totals.get('办公室', 100) * 100) if resource_totals.get('办公室', 100) > 0 else 0.0,
        '卡位': (resource_stats['卡位'] / resource_totals.get('卡位', 100) * 100) if resource_totals.get('卡位', 100) > 0 else 0.0,
        '实验室': (resource_stats['实验室'] / resource_totals.get('实验室', 100) * 100) if resource_totals.get('实验室', 100) > 0 else 0.0,
        '负80': None,  # 负80冰箱不计算占比
    }
    
    # 6. 本月异常条数
    exception_count = month_data[month_data['异常'].notna() & (month_data['异常'] != '')].shape[0]
    total_count = month_data.shape[0]
    exception_rate = (exception_count / total_count * 100) if total_count > 0 else 0.0
    
    return {
        '本月总费用': round(total_fee, 2),
        '本月各中心费用': {k: round(v, 2) for k, v in center_fees.items()},
        '本月TOP1课题组费用': {'课题组': top1_team, '金额': round(top1_amount, 2)},
        '本月实验室费用占比': round(lab_ratio, 2),
        '本月在用资源数': resource_stats,
        '本月资源占用率': {k: round(v, 2) if v is not None else None for k, v in resource_usage_rates.items()},
        '本月异常条数': exception_count,
        '本月异常率': round(exception_rate, 2),
    }

def prepare_chart_data(summary_df: pd.DataFrame, target_month: str) -> Dict[str, pd.DataFrame]:
    """
    准备5种图表的数据源
    
    注意：如果目标月份是当前月份（还没过完），趋势图只显示到上个月，不包含当前月
    """
    from datetime import date
    today = date.today()
    target_year, target_month_num = map(int, target_month.split('-'))
    is_current_month = (target_year == today.year and target_month_num == today.month)
    
    # 筛选目标月份的数据
    month_data = summary_df[summary_df['年月'] == target_month].copy()
    
    # A. 费用结构（饼图/环形图）：资源类型 × 当月金额
    if not month_data.empty:
        fee_structure = month_data.groupby('资源类型')['当月金额'].sum().reset_index()
        fee_structure.columns = ['资源类型', '金额']
        fee_structure = fee_structure.sort_values('金额', ascending=False)
    else:
        fee_structure = pd.DataFrame(columns=['资源类型', '金额'])
    
    # B. 费用排名（横向条形图）：课题组TOP10 × 当月金额
    if not month_data.empty:
        team_ranking = month_data.groupby('所属课题组')['当月金额'].sum().sort_values(ascending=False).head(10).reset_index()
        team_ranking.columns = ['课题组', '当月金额']
    else:
        team_ranking = pd.DataFrame(columns=['课题组', '当月金额'])
    
    # 计算最近3个月合计（如果数据足够）
    all_months = sorted(summary_df['年月'].unique())
    target_idx = all_months.index(target_month) if target_month in all_months else -1
    if target_idx >= 2 and not month_data.empty:
        recent_3_months = all_months[target_idx - 2:target_idx + 1]
        recent_data = summary_df[summary_df['年月'].isin(recent_3_months)]
        if not recent_data.empty:
            team_3month = recent_data.groupby('所属课题组')['当月金额'].sum().sort_values(ascending=False).head(10).reset_index()
            team_3month.columns = ['课题组', '最近3个月合计']
            team_ranking = team_ranking.merge(team_3month, on='课题组', how='left').fillna(0)
        else:
            team_ranking['最近3个月合计'] = team_ranking['当月金额']
    else:
        if not team_ranking.empty:
            team_ranking['最近3个月合计'] = team_ranking['当月金额']
        else:
            team_ranking['最近3个月合计'] = pd.Series(dtype=float)
    
    # C. 中心对比（堆叠柱状）：所属中心 × 资源类型 × 当月金额
    if not month_data.empty:
        center_comparison = month_data.groupby(['所属中心', '资源类型'])['当月金额'].sum().reset_index()
        center_comparison.columns = ['所属中心', '资源类型', '金额']
    else:
        center_comparison = pd.DataFrame(columns=['所属中心', '资源类型', '金额'])
    
    # D. 月度趋势（折线）：年月 × 总费用
    # 如果目标月份是当前月份（还没过完），趋势只显示到上个月，不包含当前月
    if not summary_df.empty:
        trend_df = summary_df.copy()
        if is_current_month:
            # 排除当前月份，只显示到上个月
            trend_df = trend_df[trend_df['年月'] != target_month]
        
        if not trend_df.empty:
            monthly_trend = trend_df.groupby('年月')['当月金额'].sum().reset_index()
            monthly_trend.columns = ['年月', '总费用']
            
            # 按中心分组的趋势
            monthly_trend_by_center = trend_df.groupby(['年月', '所属中心'])['当月金额'].sum().reset_index()
            monthly_trend_by_center.columns = ['年月', '所属中心', '总费用']
            
            # 按课题组分组的趋势（TOP5课题组）
            if not month_data.empty:
                top_teams = month_data.groupby('所属课题组')['当月金额'].sum().sort_values(ascending=False).head(5).index.tolist()
                if top_teams:
                    monthly_trend_by_team = trend_df[trend_df['所属课题组'].isin(top_teams)].groupby(['年月', '所属课题组'])['当月金额'].sum().reset_index()
                    monthly_trend_by_team.columns = ['年月', '所属课题组', '总费用']
                else:
                    monthly_trend_by_team = pd.DataFrame(columns=['年月', '所属课题组', '总费用'])
            else:
                monthly_trend_by_team = pd.DataFrame(columns=['年月', '所属课题组', '总费用'])
        else:
            monthly_trend = pd.DataFrame(columns=['年月', '总费用'])
            monthly_trend_by_center = pd.DataFrame(columns=['年月', '所属中心', '总费用'])
            monthly_trend_by_team = pd.DataFrame(columns=['年月', '所属课题组', '总费用'])
    else:
        monthly_trend = pd.DataFrame(columns=['年月', '总费用'])
        monthly_trend_by_center = pd.DataFrame(columns=['年月', '所属中心', '总费用'])
        monthly_trend_by_team = pd.DataFrame(columns=['年月', '所属课题组', '总费用'])
    
    # E. 资源清单（Top资源卡片）：当月金额最高的资源TOP10
    if not month_data.empty and len(month_data) > 0:
        resource_list = month_data.nlargest(10, '当月金额')[['所属课题组', '资源类型', '名称', '编号', '当月金额', '备注']].copy()
        resource_list.columns = ['课题组', '资源类型', '名称', '编号', '当月金额', '备注']
    else:
        resource_list = pd.DataFrame(columns=['课题组', '资源类型', '名称', '编号', '当月金额', '备注'])
    
    return {
        '费用结构': fee_structure,
        '费用排名': team_ranking,
        '中心对比': center_comparison,
        '月度趋势': monthly_trend,
        '月度趋势_按中心': monthly_trend_by_center,
        '月度趋势_按课题组': monthly_trend_by_team,
        '资源清单': resource_list,
    }

def prepare_table_data(summary_df: pd.DataFrame, target_month: str) -> Dict[str, pd.DataFrame]:
    """
    准备2种表格数据
    
    环比计算逻辑：
    - 如果目标月份是当前月份（还没过完）：环比 = (上月 - 上上月) / 上上月 × 100
      因为本月数据不完整，所以用上月和上上月（都是完整月份）进行比较
    - 如果目标月份是历史月份（已过完）：环比 = (本月 - 上月) / 上月 × 100
      正常计算本月相对于上月的增长率
    """
    from datetime import date
    today = date.today()
    target_year, target_month_num = map(int, target_month.split('-'))
    is_current_month = (target_year == today.year and target_month_num == today.month)
    
    # 表1：课题组费用表（前20）
    # 列：所属课题组 | 所属中心 | 本月金额 | 上月金额 | 环比% | 累计金额(选中范围)
    
    all_months = sorted(summary_df['年月'].unique()) if not summary_df.empty else []
    target_idx = all_months.index(target_month) if target_month in all_months else -1
    
    # 本月数据
    month_data = summary_df[summary_df['年月'] == target_month].copy()
    if not month_data.empty:
        team_month = month_data.groupby(['所属课题组', '所属中心'])['当月金额'].sum().reset_index()
        team_month.columns = ['所属课题组', '所属中心', '本月金额']
    else:
        team_month = pd.DataFrame(columns=['所属课题组', '所属中心', '本月金额'])
    
    # 上月数据
    if not team_month.empty:
        if target_idx > 0:
            prev_month = all_months[target_idx - 1]
            prev_data = summary_df[summary_df['年月'] == prev_month].copy()
            if not prev_data.empty:
                team_prev = prev_data.groupby('所属课题组')['当月金额'].sum().reset_index()
                team_prev.columns = ['所属课题组', '上月金额']
                team_month = team_month.merge(team_prev, on='所属课题组', how='left').fillna(0)
            else:
                team_month['上月金额'] = 0.0
        else:
            team_month['上月金额'] = 0.0
        
        # 计算环比
        # 如果当前月份还没过完：环比 = (上月 - 上上月) / 上上月 × 100（因为本月数据不完整，用上月和上上月比较）
        # 如果历史月份（已过完）：环比 = (本月 - 上月) / 上月 × 100
        if is_current_month and target_idx > 1:
            # 当前月份还没过完，计算上个月相对于上上个月的环比
            prev_prev_month = all_months[target_idx - 2]
            prev_prev_data = summary_df[summary_df['年月'] == prev_prev_month].copy()
            if not prev_prev_data.empty:
                team_prev_prev = prev_prev_data.groupby('所属课题组')['当月金额'].sum().reset_index()
                team_prev_prev.columns = ['所属课题组', '上上月金额']
                team_month = team_month.merge(team_prev_prev, on='所属课题组', how='left').fillna(0)
                # 环比 = (上月 - 上上月) / 上上月 × 100
                team_month['环比%'] = team_month.apply(
                    lambda row: ((row['上月金额'] - row['上上月金额']) / row['上上月金额'] * 100) if row['上上月金额'] > 0 else 0.0,
                    axis=1
                )
            else:
                team_month['环比%'] = 0.0
        else:
            # 历史月份（已过完），计算本月相对于上月的环比
            team_month['环比%'] = team_month.apply(
                lambda row: ((row['本月金额'] - row['上月金额']) / row['上月金额'] * 100) if row['上月金额'] > 0 else 0.0,
                axis=1
            )
        team_month['环比%'] = team_month['环比%'].round(2)
        
        # 根据是否当前月份，动态设置环比列名
        if is_current_month and target_idx > 1:
            # 当前月份：环比表示上月vs上上月
            team_month = team_month.rename(columns={'环比%': '环比%（上月vs上上月）'})
        else:
            # 历史月份：环比表示本月vs上月
            team_month = team_month.rename(columns={'环比%': '环比%（本月vs上月）'})
        
        # 累计金额（从最早月份到目标月份）
        if target_idx >= 0:
            cumulative_months = all_months[:target_idx + 1]
            cumulative_data = summary_df[summary_df['年月'].isin(cumulative_months)]
            if not cumulative_data.empty:
                team_cumulative = cumulative_data.groupby('所属课题组')['当月金额'].sum().reset_index()
                team_cumulative.columns = ['所属课题组', '累计金额']
                team_month = team_month.merge(team_cumulative, on='所属课题组', how='left').fillna(0)
            else:
                team_month['累计金额'] = team_month['本月金额']
        else:
            team_month['累计金额'] = team_month['本月金额']
        
        # 确定环比列名（根据是否当前月份）
        if is_current_month and target_idx > 1:
            mom_col_name = '环比%（上月vs上上月）'
        else:
            mom_col_name = '环比%（本月vs上月）'
        
        # 排序并取前20
        team_month = team_month.sort_values('本月金额', ascending=False).head(20)
        team_month = team_month[['所属课题组', '所属中心', '本月金额', '上月金额', mom_col_name, '累计金额']]
    else:
        # 确定默认列名
        if is_current_month:
            mom_col_name = '环比%（上月vs上上月）'
        else:
            mom_col_name = '环比%（本月vs上月）'
        team_month = pd.DataFrame(columns=['所属课题组', '所属中心', '本月金额', '上月金额', mom_col_name, '累计金额'])
    
    # 表2：异常列表（最多10条）
    month_data_with_exception = summary_df[
        (summary_df['年月'] == target_month) & 
        (summary_df['异常'].notna()) & 
        (summary_df['异常'] != '')
    ].copy()
    
    if not month_data_with_exception.empty:
        exception_list = month_data_with_exception[['所属课题组', '资源类型', '编号', '异常', '年月']].copy()
        exception_list.columns = ['所属课题组', '资源类型', '编号', '异常内容', '涉及月份']
        exception_list = exception_list.head(10)
    else:
        exception_list = pd.DataFrame(columns=['所属课题组', '资源类型', '编号', '异常内容', '涉及月份'])
    
    return {
        '课题组费用表': team_month,
        '异常列表': exception_list,
    }

def export_to_json(kpis: Dict[str, Any], chart_data: Dict[str, pd.DataFrame], 
                   table_data: Dict[str, pd.DataFrame], target_month: str, 
                   output_json_path: str) -> None:
    """
    导出数据为JSON格式，供Web前端使用
    
    参数:
        kpis: KPI指标数据
        chart_data: 图表数据
        table_data: 表格数据
        target_month: 目标月份
        output_json_path: 输出JSON文件路径
    """
    import json
    
    # 处理资源占用率，过滤掉None值（负80）
    resource_usage_rates = {}
    for k, v in kpis['本月资源占用率'].items():
        if v is not None:
            resource_usage_rates[k] = v
    
    result = {
        'meta': {
            '生成时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            '目标月份': target_month,
        },
        'kpis': {
            '本月总费用': float(kpis['本月总费用']),
            '本月TOP1课题组': {
                '课题组': kpis['本月TOP1课题组费用']['课题组'],
                '金额': float(kpis['本月TOP1课题组费用']['金额'])
            },
            '本月实验室费用占比': float(kpis['本月实验室费用占比']),
            '本月在用资源数': {
                '办公室': int(kpis['本月在用资源数']['办公室']),
                '卡位': int(kpis['本月在用资源数']['卡位']),
                '实验室': int(kpis['本月在用资源数']['实验室']),
                '负80': int(kpis['本月在用资源数']['负80']),
            },
            '本月资源占用率': resource_usage_rates,
            '本月各中心费用': {k: float(v) for k, v in kpis['本月各中心费用'].items()},
        },
        'charts': {
            '费用结构': chart_data['费用结构'].to_dict('records'),
            '费用排名': chart_data['费用排名'].to_dict('records'),
            '中心对比': chart_data['中心对比'].to_dict('records'),
            '月度趋势': chart_data['月度趋势'].to_dict('records'),
            '月度趋势_按中心': chart_data['月度趋势_按中心'].to_dict('records'),
            '月度趋势_按课题组': chart_data['月度趋势_按课题组'].to_dict('records'),
            '资源清单': chart_data['资源清单'].to_dict('records'),
        },
        'tables': {
            '课题组费用表': table_data['课题组费用表'].to_dict('records'),
            '异常列表': table_data['异常列表'].to_dict('records'),
        }
    }
    
    # 处理NaN值，转换为null
    def convert_nan(obj):
        if isinstance(obj, dict):
            return {k: convert_nan(v) for k, v in obj.items()}
        elif isinstance(obj, list):
            return [convert_nan(item) for item in obj]
        elif isinstance(obj, float) and pd.isna(obj):
            return None
        elif isinstance(obj, (pd.Timestamp, datetime)):
            return obj.strftime('%Y-%m-%d %H:%M:%S')
        else:
            return obj
    
    result = convert_nan(result)
    
    with open(output_json_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2, default=str)