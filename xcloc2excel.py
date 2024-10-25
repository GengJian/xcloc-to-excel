import os
import tkinter as tk
import xml.etree.ElementTree as ET
from tkinter import messagebox
from openpyxl import Workbook

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

    # 冻结第一行
    ws.freeze_panes = 'A2'  # 这里的 'A2' 表示冻结第一行，第二行开始是活动行
    # 设置每列的宽度为自适应宽度
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter  # 获取列字母
        for cell in column:
            try:
                if cell.value is not None:  # 确保单元格不为空
                    cell_value = str(cell.value)
                    max_length = max(max_length, len(cell_value))  # 更新最大长度
            except Exception as e:
                print(f"调整列宽时发生错误: {e}")
        adjusted_width = (max_length + 2)  # 可以加一点额外的宽度
        ws.column_dimensions[column_letter].width = adjusted_width
    # 保存Excel文件
    wb.save(output_excel)
    print(f"转换完成，Excel文件已保存至: {output_excel}")

def run_conversion():
    xcloc_folder = folder_entry.get()
    output_excel = output_entry.get()
    
    if not os.path.isdir(xcloc_folder):
        messagebox.showerror("错误", "无效的文件夹路径！")
        return
        
    if not output_excel.endswith('.xlsx'):
        messagebox.showerror("错误", "输出文件名必须以 .xlsx 结尾！")
        return

    convert_xcloc_to_excel(xcloc_folder, output_excel)
    messagebox.showinfo("完成", "转换已完成！")

# 创建图形界面
root = tk.Tk()
root.title("XLIFF to Excel Converter")

# 创建输入框和标签
tk.Label(root, text="输入 Xcloc 文件的完整路径:").pack(pady=5)
folder_entry = tk.Entry(root, width=50)
folder_entry.pack(pady=5)

tk.Label(root, text="输出 Excel 文件名/完整目录 (.xlsx):").pack(pady=5)
output_entry = tk.Entry(root, width=50)
output_entry.pack(pady=5)

# 创建按钮
convert_button = tk.Button(root, text="转换", command=run_conversion)
convert_button.pack(pady=20)

# 运行主循环
root.mainloop()