from typing import List
from datetime import timedelta
import pandas as pd
import re
import sys
import argparse
import json
import os
import math


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
    # 重命名列
    df_etl = df_etl.rename(columns={
        '环缝':'连接管根数',
        '中频弯': '线弯',
        '冷弯': '单机弯'
    })
    # df_etl['炉型']=df_etl['WBS元素'].apply(extract_letter_number_code)
    # df_etl.loc[:, '炉型'] = df_etl['WBS元素'].apply(extract_letter_number_code)
    df_etl = df_etl.assign(炉型=df_etl['WBS元素'].apply(extract_letter_number_code))
    df_etl=df_etl.sort_values(by=['WBS元素','实际投料日期'],ascending=[True,True])
    return df_etl


def get_unique_filename(base_path, filename):
    # 获取文件名和扩展名
    name, ext = os.path.splitext(filename)
    counter = 1
    
    # 构建完整路径
    directory = os.path.dirname(base_path)  # 获取上级目录
    full_path = os.path.join(directory, filename)
    
    # 如果文件已存在，添加(1)、(2)等后缀
    while os.path.exists(full_path):
        new_filename = f"{name}({counter}){ext}"
        full_path = os.path.join(directory, new_filename)
        counter += 1
    
    return full_path


class guolu_opt():
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
        df = df.reset_index(drop=True)
        for index, row in df.iterrows():
            subcontractor = row['预估单位']
            if subcontractor and subcontractor in self.manufacturing_ability:
                first_key = next(iter(self.manufacturing_ability[subcontractor])) # 简化了任务

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
                    else:
                        base_date = df.at[index-1, 'AI推算完工时间']

                    completion_date = base_date + timedelta(days=math.ceil(extra_days))

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


if __name__ == "__main__":
    # 确保 stdout 使用 UTF-8 编码
    if sys.stdout.encoding != "utf-8":
        sys.stdout.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(
        description="Boiler plant production scheduling optimization"
    )
    parser.add_argument(
        "--file_path", type=str, required=True, help="Path to the opt Excel file"
    )
    args = parser.parse_args()

    guolu_opt = guolu_opt()
    result_df = guolu_opt.process_sheet(args.file_path)

    # 将 result_dict 转换为 json 字符串
    # result_dict_json = json.dumps(result_dict, indent=4, ensure_ascii=False)
    # print(result_dict_json)
    # result_df.to_excel('result.xlsx', index=False)

    try:
        # 尝试保存文件
        output_path = get_unique_filename(args.file_path, 'opt_result.xlsx')
        result_df.to_excel(output_path, index=False)
        print(f"锅炉排程优化结果已保存到: {output_path}")
    except Exception as e:
        print(f"保存文件时出错: {str(e)}")