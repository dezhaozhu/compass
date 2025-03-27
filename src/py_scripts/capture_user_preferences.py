import argparse
import json
import os
import sys
from typing import Dict

import pandas as pd


def capture_user_preferences(
    file_path: str, preferences_path: str
) -> Dict[str, Dict[str, int]]:
    """这是一个捕捉历史排程中人工排程偏好的工具, 根据输入的文件路径，返回人工偏好信息
    人工偏好信息主要包括每个产品的工厂优先级
    Args:
        file_path: str, 文件路径
        preferences_path: str, 人工偏好信息保存路径
    Returns:
        dict[str, dict[str, int]], 人工偏好信息。
    """
    df = pd.read_excel(file_path)
    # 取 "令号" 列字符串的第四到第7个字符
    df["spec_令号"] = df["令号"].str.slice(6, 9)

    # 移除 "分包单位" 列 中包含 ；的行
    df = df[~df["分包单位"].str.contains("；")]

    # 计算每个spec_令号下各分包单位的数量
    grouped = df.groupby(["spec_令号", "分包单位"]).size().reset_index(name="数量")

    # 计算每个spec_令号的总数
    total_counts = df.groupby("spec_令号").size().to_dict()

    # 计算占比并构建结果字典
    result_dict = {}

    for spec, group_df in grouped.groupby("spec_令号"):
        if spec not in result_dict:
            result_dict[spec] = {}

        total = total_counts[spec]

        for _, row in group_df.iterrows():
            unit = row["分包单位"]
            count = row["数量"]
            proportion = count / total
            result_dict[spec][unit] = round(proportion, 2)  # 保留两位小数

    # 对 result_dict 中每个 值的 字典 进行排序
    result_dict = {
        k: dict(sorted(v.items(), key=lambda x: x[1], reverse=True))
        for k, v in result_dict.items()
    }
    # target_path = os.path.join(str(preferences_path), "preferences.json")
    # print("target_path", target_path)
    # 把 result_dict 写入到 json 文件中, 如果文件不存在, 则创建文件
    with open(str(preferences_path), "w", encoding="utf-8") as f:
        json.dump(result_dict, f, ensure_ascii=False)

    return result_dict


if __name__ == "__main__":
    # 确保 stdout 使用 UTF-8 编码
    if sys.stdout.encoding != "utf-8":
        sys.stdout.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(
        description="Capture user preferences from Excel file"
    )
    parser.add_argument(
        "--file_path", type=str, required=True, help="Path to the Excel file"
    )
    parser.add_argument(
        "--preferences_path", type=str, required=True, help="Path to save preferences"
    )
    args = parser.parse_args()

    result_dict = capture_user_preferences(args.file_path, args.preferences_path)

    # 将 result_dict 转换为 json 字符串
    result_dict_json = json.dumps(result_dict, indent=4, ensure_ascii=False)
    print(result_dict_json)
