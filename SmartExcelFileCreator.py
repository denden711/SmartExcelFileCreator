import os
from openpyxl import Workbook
import tkinter as tk
from tkinter import filedialog

# GUIウィンドウの初期化を行い、ウィンドウは表示しません。
root = tk.Tk()
root.withdraw()

# ファイルダイアログを開いてフォルダのパスを選択させます。
folder_path = filedialog.askdirectory(title="フォルダを選択してください")

# フォルダ内のすべてのファイル名を取得します。
file_names = os.listdir(folder_path)

def create_excel_file(base_name):
    # Excelファイルのフルパスを生成します。
    excel_path = os.path.join(folder_path, f"{base_name}.xlsx")
    
    # 既に同名のExcelファイルが存在するかチェックします。
    if not os.path.exists(excel_path):
        # 新しいワークブック（Excelファイル）を作成します。
        wb = Workbook()
        # Excelファイルを保存します。ファイル名は元のファイル名に基づきますが、拡張子は.xlsxになります。
        wb.save(excel_path)

for file_name in file_names:
    # ファイル名から拡張子を除いた基本名を取得します。
    base_name, ext = os.path.splitext(file_name)
    # ファイルがExcelファイル自体でないことを確認します。
    if ext.lower() != '.xlsx':
        # 同じ基本名でExcelファイルを作成しますが、既に存在する場合はスキップします。
        create_excel_file(base_name)
