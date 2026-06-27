import win32com.client as win32
import os
import shutil


def update_add_share_xlsm():
    """
    使用 Excel 应用程序更新 add share 列
    流程：打开源文件 → 创建备份 → 计算并写入
    计算公式：(本行share - 上一行share) / 10000
    备份文件：2026年QDII LOF基金估值偏差记录_backup.xlsm
    """

    file_path = r"E:\2026年QDII LOF基金估值偏差记录.xlsm"
    backup_path = r"E:\2026年QDII LOF基金估值偏差记录_backup.xlsm"

    print("=" * 60)
    print("更新 add share 列")
    print("计算公式：(本行 - 上一行) / 10000")
    print("=" * 60)

    if not os.path.exists(file_path):
        print(f"❌ 文件不存在：{file_path}")
        return

    # 1. 创建备份
    shutil.copy2(file_path, backup_path)
    print(f"✅ 备份已创建：{os.path.basename(backup_path)}")

    # 2. 启动 Excel
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    print("📖 打开源文件...")
    wb = excel.Workbooks.Open(file_path)

    # 3. 处理每个工作表
    for sheet in wb.Sheets:
        print(f"\n📋 工作表：{sheet.Name}")

        share_col = None
        add_share_col = None

        for col in range(1, sheet.UsedRange.Columns.Count + 1):
            val = sheet.Cells(1, col).Value
            if val:
                val_str = str(val).strip().lower()
                if val_str == 'share':
                    share_col = col
                elif val_str == 'add share':
                    add_share_col = col

        if share_col is None or add_share_col is None:
            print(f"  ⚠️ 未找到列，跳过")
            continue

        print(f"  share 列：{share_col}")
        print(f"  add share 列：{add_share_col}")

        last_row = sheet.Cells(sheet.Rows.Count, share_col).End(-4162).Row

        count = 0
        for row in range(3, last_row + 1):
            curr = sheet.Cells(row, share_col).Value
            prev = sheet.Cells(row - 1, share_col).Value

            if curr is not None and prev is not None:
                try:
                    result = (float(curr) - float(prev)) / 10000
                    sheet.Cells(row, add_share_col).Value = result
                    count += 1
                except:
                    pass

        print(f"  ✅ 更新 {count} 行")

    # 4. 保存
    print("\n💾 保存源文件...")
    wb.Save()
    wb.Close()
    excel.Quit()

    print("\n" + "=" * 60)
    print("✅ 完成！")
    print(f"📁 源文件：{os.path.basename(file_path)}")
    print(f"💾 备份文件：{os.path.basename(backup_path)}")
    print("=" * 60)


if __name__ == "__main__":
    update_add_share_xlsm()