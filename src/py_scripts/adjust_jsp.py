import datetime
import json
from typing import Dict, List

import numpy as np
import pandas as pd


def adjust_schedule(file_path: str, product_to_date_order: List[Dict[str, str]]):
    """这是一个调整排程工具，根据提供的调整要求，做完调整之后，会发送调整成功的消息。
    args:
        file_path: str, 需要调整的文件路径
        product_to_date_order: list[dict[str, str]], 包含产品单号(product_id)和日期(target_date) 和 工厂(target_factory)。
        可能需要调整的单号为多个，每个单号包含一个目标日期和工厂。
        工厂是可选的，如果工厂为空，则表示放到当前工厂的指定时间。 所有可能的厂为：上锅，汇能，绿叶,金科尔,德海, 申港, 祥安,都江电力等
        时间的格式为：2024-12-12 00:00:00，如果没详细到具体的天, 那么就放到当月第一天
        例如：[{product_id: "Z-1400D52-36-560", target_date: "2024-12-12 00:00:00", target_factory: "上锅"}, {product_id: "Z-1400F15-259-740(5)", target_date: "2024-12-12 00:00:00", target_factory: "汇能"}] 表示需要调整 Z-1400D52-36-560 和 Z-1400F15-259-740(5) 的日期为 2024-12-12 00:00:00。
    """
    # 如果指定了 工厂 和 时间，那么 直接插入到 指定工厂 的 指定时间.
    # 如果仅仅指定了时间，那么 插入到当前工厂的 指定时间
    df = pd.read_excel(file_path)
    # 根据 product_to_date_order 中的 product_id 和 target_date 调整 df 中的日期
    info = []
    for order in product_to_date_order:
        try:
            product_id = order["product_id"]
            target_date = order["target_date"]

            cur_data = df.loc[df["WBS元素"] == product_id, "AI推算开始时间"].values[0]

            # 'numpy.datetime64' 转为 datetime.datetime
            if not np.isnan(cur_data):
                cur_data = pd.Timestamp(cur_data).to_pydatetime()
            else:
                cur_data = None

            # 当前日期为 None 表示没有排程日期， 直接填写
            if cur_data is None:
                df.loc[df["WBS元素"] == product_id, "AI推算开始时间"] = target_date
                df.loc[df["WBS元素"] == product_id, "预估单位"] = order[
                    "target_factory"
                ]
                info.append(f"单号 {product_id} 没有排程日期，直接填写")
                continue

            target_date = datetime.datetime.strptime(target_date, "%Y-%m-%d %H:%M:%S")

            # 如果 target_date 和当前时间相差不超过1天，直接返回
            if (target_date - cur_data).days <= 1:
                info.append(
                    f"单号 {product_id} 的排程日期和当前时间相差不超过1天，不需要调整"
                )
                continue
            # 如果 target_date 和当前时间相差超过5天，那么 找到 指定  target_factory 最靠近 target_date 的日期
            start_date = (
                order["target_factory"]
                if order["target_factory"] is not None
                else df.loc[df["WBS元素"] == product_id, "预估单位"].values[0]
            )
            # 修改 product_id 的工厂和日期。
            df.loc[df["WBS元素"] == product_id, "AI推算开始时间"] = target_date
            df.loc[df["WBS元素"] == product_id, "预估单位"] = start_date
            info.append(f"单号 {product_id} 调整排程成功")
        except Exception as e:
            info.append(f"单号 {product_id} 调整排程失败: {e}")
            continue

    try:
        # 保存调整后的 df 到 csv 中
        df.to_excel(file_path, index=False)
    except Exception as e:
        info.append(f"保存调整后的 df 到 csv 中失败: {e}")

    return "\n".join(info)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Adjust JSP schedule")
    parser.add_argument(
        "--file_path", type=str, required=True, help="Path to the JSP file"
    )
    parser.add_argument(
        "--product_to_date_order", type=str, required=True, help="Product to date order"
    )
    args = parser.parse_args()

    try:
        file_path = args.file_path
        product_to_date_order = json.loads(args.product_to_date_order)
        # file_path = "./opt_result.xlsx"
        # product_to_date_order = [
        #     {
        #         "product_id": "Z-1400D52-35-560",
        #         "target_date": "2024-01-14 00:00:00",
        #         "target_factory": "上锅",
        #     },
        #     {
        #         "product_id": "Z-1400F15-260-740(7)",
        #         "target_date": "2024-12-02 00:00:00",
        #         "target_factory": "汇能",
        #     },
        # ]
        result = adjust_schedule(file_path, product_to_date_order)
        print(result)
    except Exception as e:
        print(f"调整排程失败: {e}")
