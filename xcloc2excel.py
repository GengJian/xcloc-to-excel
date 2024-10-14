import os
import xml.etree.ElementTree as ET
from openpyxl import Workbook
import argparse

def convert_xcloc_to_excel(xcloc_folder, output_excel):
    # 创建一个 Excel Workbook
    wb = Workbook()
    ws = wb.active
    ws.append(["Key", "Source", "Target"])  # 添加标题行，包括Key

    # 处理指定文件夹中的所有XLIFF文件
    for root, dirs, files in os.walk(xcloc_folder):
        found_xliff_file = False  # 标记是否找到XLIFF文件
        for file in files:
            if file.endswith(".xliff"):
                found_xliff_file = True
                xliff_path = os.path.join(root, file)

                # 尝试解析XLIFF文件
                try:
                    tree = ET.parse(xliff_path)
                    root_element = tree.getroot()

                    # 使用XPath查找所有的trans-unit
                    for trans_unit in root_element.findall('.//{urn:oasis:names:tc:xliff:document:1.2}trans-unit'):
                        key = trans_unit.get('id')  # 获取key
                        source = trans_unit.find('{urn:oasis:names:tc:xliff:document:1.2}source')
                        target = trans_unit.find('{urn:oasis:names:tc:xliff:document:1.2}target')

                        # 提取文本
                        source_text = source.text if source is not None else ''
                        target_text = target.text if target is not None else ''

                        # 将数据写入Excel
                        ws.append([key, source_text, target_text])
                except Exception as e:
                    print(f"解析文件 {xliff_path} 失败: {e}")

        if not found_xliff_file:
            print(f"在目录 {xcloc_folder} 中没有找到任何XLIFF文件。")

    # 保存Excel文件
    wb.save(output_excel)
    print(f"转换完成，Excel文件已保存至: {output_excel}")

def main():
    parser = argparse.ArgumentParser(description='Convert Xcode .xcloc files to Excel.')
    parser.add_argument('xcloc_folder', help='The path to the .xcloc folder')
    parser.add_argument('output_excel', help='The path for the output Excel file')

    args = parser.parse_args()

    convert_xcloc_to_excel(args.xcloc_folder, args.output_excel)

if __name__ == "__main__":
    main()