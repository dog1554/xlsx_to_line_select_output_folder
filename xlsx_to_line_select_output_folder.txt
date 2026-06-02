import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd


def clean_cell(cell):
    """将单元格内容转成字符串，并移除所有空白字符。"""
    if pd.isna(cell):
        return None
    return re.sub(r"\s+", "", str(cell))


def get_available_output_path(output_folder, original_filename):
    """
    根据原始 xlsx 文件名，在指定输出文件夹中生成 .line 路径。
    如果同名文件已存在，返回 None，避免覆盖。
    """
    base_name = os.path.splitext(os.path.basename(original_filename))[0]
    line_filename = base_name + ".line"
    line_path = os.path.join(output_folder, line_filename)

    if os.path.exists(line_path):
        return None, line_filename

    return line_path, line_filename


def convert_xlsx_to_line(xlsx_path, output_folder):
    """读取单个 xlsx 文件，并在手动选择的输出文件夹中生成同名 .line 文件。"""
    filename = os.path.basename(xlsx_path)

    # 跳过 Excel 临时文件
    if filename.startswith("~$"):
        return "skipped", f"{filename}（临时文件，已跳过）"

    if not filename.lower().endswith(".xlsx"):
        return "skipped", f"{filename}（不是 .xlsx 文件，已跳过）"

    line_path, line_filename = get_available_output_path(output_folder, filename)

    # 避免覆盖已有 .line 文件
    if line_path is None:
        return "skipped", f"{filename} → {line_filename}（输出文件夹中目标文件已存在，已跳过）"

    try:
        # header=None：不把第一行当作表头
        df = pd.read_excel(xlsx_path, header=None)

        with open(line_path, "w", encoding="utf-8") as f:
            for _, row in df.iterrows():
                cleaned_cells = []
                for cell in row:
                    cleaned_cell = clean_cell(cell)
                    if cleaned_cell is not None:
                        cleaned_cells.append(cleaned_cell)

                # 用英文逗号连接每一行的所有非空单元格
                f.write(",".join(cleaned_cells) + "\n")

        return "success", f"{filename} → {line_filename}"

    except Exception as e:
        return "error", f"{filename} 转换失败：{e}"


def select_files_and_output_folder_then_convert():
    # 隐藏主窗口，只显示选择窗口
    root = tk.Tk()
    root.withdraw()

    # 第一步：手动选择要转换的 xlsx 文件，可多选
    xlsx_files = filedialog.askopenfilenames(
        title="请选择要转换的 .xlsx 文件（可多选）",
        filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
    )

    if not xlsx_files:
        root.destroy()
        return

    # 第二步：手动选择 .line 输出文件夹
    output_folder = filedialog.askdirectory(title="请选择 .line 文件输出文件夹")

    if not output_folder:
        root.destroy()
        return

    success_list = []
    skipped_list = []
    error_list = []

    for xlsx_path in xlsx_files:
        status, message = convert_xlsx_to_line(xlsx_path, output_folder)
        if status == "success":
            success_list.append(message)
        elif status == "skipped":
            skipped_list.append(message)
        else:
            error_list.append(message)

    msg = (
        f"完成！\n"
        f"输出文件夹：{output_folder}\n"
        f"成功生成 {len(success_list)} 个 .line 文件"
    )

    if skipped_list:
        msg += f"\n跳过: {len(skipped_list)}"
    if error_list:
        msg += f"\n错误: {len(error_list)}"

    details = success_list + skipped_list + error_list
    if details:
        msg += "\n\n详情：\n" + "\n".join(details[:10])
        if len(details) > 10:
            msg += f"\n……还有 {len(details) - 10} 条未显示"

    messagebox.showinfo("xlsx 转 line 完成", msg)
    root.destroy()


if __name__ == "__main__":
    select_files_and_output_folder_then_convert()
