from typing import List
from datetime import timedelta
import pandas as pd
import re


def extract_letter_number_code(text):
    # 匹配模式：数字后面跟着一个字母和数字，后面是-或.
    # 这里使用两种模式：一种是字母后跟数字再跟-，另一种是字母后跟数字再跟.
    pattern = r'(\d+)([A-Za-z]\d+)[-.]'
    
    # 搜索匹配
    match = re.search(pattern, text)
    
    if match:
        # 匹配到了字母+数字的组合
        letter_number_code = match.group(2)  # 提取第二个捕获组，即字母+数字部分
        return letter_number_code
    else:
        return None


def etl_data(file_path):
    df=pd.read_excel(file_path)
    df_etl=df[['WBS元素','项目名称','部件号','部件名称','实际投料日期','计划完工时间','预估单位','重量吨','数量支','环缝','中频弯','冷弯','技术需求','参照炉型','备注']]
    df_etl['炉型']=df_etl['WBS元素'].apply(extract_letter_number_code)
    df_etl=df_etl.sort_values(by=['WBS元素','实际投料日期'],ascending=[True,True])
    return df_etl


class ExcelToolkit():
    def __init__(self):
        self.manufacturing_ability={
        "上锅": {
            "产线1": {
                "机械焊": 480,
                "线弯": 500,
                "单机弯": 1000
            },
            "产线2": {
                "机械焊": 480,
                "线弯": 500,
                "单机弯": 1000
            }
        },
        "申港": {
            "产线1": {
                "机械焊": 480,
                "线弯": 800,
                "单机弯": 500
            },
            "产线2": {
                "机械焊": 480,
                "线弯": 800,
                "单机弯": 500
            }
        },
        "绿叶": {
            "产线1": {
                "机械焊": 800,
                "线弯": 1400,
                "单机弯": 2000
            },
            "产线2": {
                "机械焊": 680,
                "线弯": 1000,
                "单机弯": 2100
            },
            "产线3": {
                "机械焊": 1200,
                "线弯": 1400,
                "单机弯": 2200
            }
        }
    }

        self.factory_totals = {}
        for factory, production_lines in self.manufacturing_ability.items():
            # 初始化该工厂的各类加工能力计数器
            self.factory_totals[factory] = {
                "机械焊": 0,
                "线弯": 0,
                "单机弯": 0
            }
            # 累加该工厂所有产线的加工能力
            for line in production_lines.values():
                self.factory_totals[factory]["机械焊"] += line["机械焊"]
                self.factory_totals[factory]["线弯"] += line["线弯"]
                self.factory_totals[factory]["单机弯"] += line["单机弯"]


    def process_sheet(self,file_path):
        df = etl_data(file_path)

        df = df.sort_values(by=['预估单位', '计划完工时间', 'WBS元素'], ascending=[True, True, True])

        df['AI推算开始时间'] = None
        df['AI推算完工时间'] = None
        df['AI推算完工情况'] = ''

        df['计划完工时间'] = pd.to_datetime(df['计划完工时间'])
        df['计划完工时间'] = df['计划完工时间'] - timedelta(days=7)

        # for index, row in df.iterrows():
        #     wbs = row['WBS元素']
        #     if wbs in human_preference:
        #         df.at[index, 'AI推算分包单位'] = human_preference[wbs]
        previous_subcontractor = None

        for index, row in df.iterrows():
            subcontractor = row['预估单位']
            if subcontractor and subcontractor in self.manufacturing_ability:
                first_key = self.manufacturing_ability[subcontractor] # 简化了任务

                line_bend_ability = self.manufacturing_ability[subcontractor][first_key]['线弯']
                single_bend_ability = self.manufacturing_ability[subcontractor][first_key]['单机弯']
                
                line_bend = row.get('线弯', 0)
                single_bend = row.get('单机弯', 0)
                try:
                    line_bend = float(line_bend) if line_bend else 0
                    single_bend = float(single_bend) if single_bend else 0
                    
                    extra_days = 0
                    if line_bend_ability > 0:
                        extra_days += line_bend / line_bend_ability
                    if single_bend_ability > 0:
                        extra_days += single_bend / single_bend_ability
                    
                    # 如果分包商不同或是第一次处理
                    if subcontractor != previous_subcontractor:
                        base_date = row['实际投料日期']

                    completion_date = base_date + timedelta(days=extra_days)

                    df.at[index, 'AI推算开始时间'] = base_date
                    df.at[index, 'AI推算完工时间'] = completion_date
                    
                    previous_subcontractor = subcontractor

                    if completion_date <= row['计划完工时间']:
                        df.at[index, 'AI推算完工情况'] = '提前完工'
                    else:
                        df.at[index, 'AI推算完工情况'] = '滞后完工'
                except (ValueError, TypeError):
                    continue
        return df