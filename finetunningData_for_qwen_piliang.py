
import pandas as pd
import json
import os

def process_multiple_excels(input_folder, output_json_file):
    # 存储所有数据的列表
    all_data = []

    # 遍历指定文件夹下的所有Excel文件
    for filename in os.listdir(input_folder):
        if filename.endswith(".xlsx"):
            file_path = os.path.join(input_folder, filename)

            # 读取 Excel 文件
            xls = pd.ExcelFile(file_path)

            # 遍历每个 sheet
            for sheet_name in xls.sheet_names:
                # 读取每个 sheet 中的数据
                df = pd.read_excel(xls, sheet_name, header=None)

                # 从第5行开始读取
                for index, row in df.iterrows():
                    if index >= 4:
                        # 读取单元格数据
                        cell_value1 = row.iloc[1] if pd.notna(row.iloc[1]) else ""
                        cell_value2 = row.iloc[2] if pd.notna(row.iloc[2]) else ""
                        cell_value3 = row.iloc[3] if pd.notna(row.iloc[3]) else ""
                        cell_value4 = row.iloc[4] if pd.notna(row.iloc[4]) else ""
                        cell_value5 = row.iloc[5] if pd.notna(row.iloc[5]) else ""

                        # 组成字符串data
                        content = f"安全控制点：{cell_value1}。检测项：{cell_value2}。检测结果：{cell_value3}。"
                        summary = f"检测问题：{cell_value4}。符合情况：{cell_value5}"
                        data = {"instruction": content, "input": "","output": summary},
                        # print(data)
                        # 添加到列表
                        all_data.append(json.dumps(data, ensure_ascii=False))

    # 将列表写入 JSON 文件
    with open(output_json_file, 'w', encoding='utf-8') as json_file:
        json_file.write("\n".join(all_data))

# 用法示例
excel_folder_path = 'E:\\MyWork\\'  # 替换成包含Excel文件的文件夹路径
output_json_file_path = 'sft_for_qwen.json'  # 替换成你想要保存的 JSON 文件路径

process_multiple_excels(excel_folder_path, output_json_file_path)

