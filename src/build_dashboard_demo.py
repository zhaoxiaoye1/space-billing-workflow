"""Build fictional cross-month inputs and dashboard data through the actual core."""
from pathlib import Path
from datetime import date
import json
import pandas as pd
from billing_core import stage1_compute_billing_detail, stage2_generate_report
from dashboard_metrics import split_by_month,calculate_kpis,prepare_chart_data,prepare_table_data,export_to_json

root=Path(__file__).resolve().parents[1]
out=root/'outputs'
out.mkdir(exist_ok=True)
total_area=200
share=0.5
occupied_area=total_area*share
full_electricity=300
rows=[]

def add_resource(kind,number,name,team,center,start,end,params):
    row={'所属中心':center,'所属课题组':team,'新增/腾退':'新增','计费开始时间':start,'计费截止时间':end,kind+'编号':number,kind+'名称':name,**params}
    rows.append(row)
    rows.append({**row,'新增/腾退':'腾退'})

add_resource('办公室','DEMO-O1','演示办公室','示例课题组 A','演示生物中心','2026-01-01','2026-02-28',{'办公室面积':20,'办公室公摊率':0.1,'办公室月度单价':10})
shared={'所属中心':'演示基因中心','所属课题组':'示例课题组 B','实验室编号':'DEMO-SHARED','实验室名称':'共享实验室','实验室公摊率':0.1,'实验室月度单价':10,'计费开始时间':'2026-01-20'}
rows.extend([
 {**shared,'新增/腾退':'新增','实验室面积':occupied_area,'实验室月度电费':full_electricity*share},
 {**shared,'新增/腾退':'腾退','计费截止时间':'2026-01-31','实验室面积':40,'实验室月度电费':full_electricity*(occupied_area-40)/total_area},
 {**shared,'新增/腾退':'腾退','计费截止时间':'2026-02-10','实验室面积':occupied_area-40,'实验室月度电费':0},
])
add_resource('实验室','DEMO-L2','演示实验室','示例课题组 C','演示生物中心','2026-01-01','2026-02-28',{'实验室面积':30,'实验室公摊率':0.1,'实验室月度单价':10,'实验室月度电费':80})
input_path=root/'examples/cross_month_input.xlsx'
pd.DataFrame(rows).to_excel(input_path,index=False)
detail=stage1_compute_billing_detail(str(input_path))
shared_detail=detail[detail['编号']=='DEMO-SHARED']
assert shared_detail['使用天数'].tolist()==[12,10]
assert shared_detail['费用'].round(2).tolist()==[500,250]
stage2_generate_report(detail,str(out/'cross_month_report.xlsx'))
monthly=split_by_month(detail,['2026-01','2026-02'])
monthly.to_csv(out/'dashboard_monthly.csv',index=False,encoding='utf-8-sig')
monthly.to_json(out/'dashboard_monthly.json',orient='records',force_ascii=False,indent=2)
for month in ['2026-01','2026-02']:
    kpi=calculate_kpis(monthly,month,{'办公室':8,'卡位':24,'实验室':6})
    charts=prepare_chart_data(monthly,month)
    tables=prepare_table_data(monthly,month)
    export_to_json(kpi,charts,tables,month,str(root/'docs'/('dashboard-'+month+'.json')))
    target=root/'docs'/('dashboard-'+month+'.json')
    data=json.loads(target.read_text(encoding='utf-8'))
    data['meta']['数据说明']='全部为虚构演示。资源占用按当月出现的唯一资源编号计数，不能解释成当月月底时点状态或面积使用率。'
    target.write_text(json.dumps(data,ensure_ascii=False,indent=2,allow_nan=False),encoding='utf-8')
print('Shared laboratory: January 500.00, February 250.00, total 750.00 yuan.')
