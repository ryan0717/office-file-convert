import os
import sys
import win32com.client


def is_temporary_office_file(filename):
    """Return True for Office lock/temp files that should not be converted."""
    basename = os.path.basename(filename)
    return basename.startswith("~$")


def convert_old_formats(base_dir=None):
    """Batch upgrade all .doc/.xls files in a directory to .docx/.xlsx.
    将目录中所有 .doc/.xls 文件批量升级为 .docx/.xlsx。

    Usage: python fix_old_files.py [directory_path]
    用法：python fix_old_files.py [目录路径]

    If no directory is specified, defaults to the script's own directory.
    未指定目录时，默认使用脚本所在目录。
    """
    if base_dir is None:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    # Legacy formats to process / 需要处理的旧格式后缀
    old_formats = {
        ".doc": ".docx",
        ".xls": ".xlsx"
    }

    print("Starting Office components, please wait...")
    print("正在启动 Office 组件，请耐心等待...")

    word = None
    excel = None

    # Walk through directories / 遍历文件夹
    for root, dirs, files in os.walk(base_dir):
        for filename in files:
            if is_temporary_office_file(filename):
                continue
            file_ext = os.path.splitext(filename)[1].lower()
            file_path = os.path.join(root, filename)

            if file_ext in old_formats:
                print(f"Found legacy file / 发现旧格式文件: {filename}, converting / 正在转换...")

                try:
                    if file_ext == ".doc":
                        # Lazy-load Word / 懒加载 Word
                        if word is None:
                            word = win32com.client.Dispatch("Word.Application")
                            word.Visible = False

                        doc = word.Documents.Open(file_path)
                        new_path = os.path.splitext(file_path)[0] + ".docx"
                        # 16 = wdFormatXMLDocument (.docx)
                        doc.SaveAs2(new_path, FileFormat=16)
                        doc.Close()
                        print(f"  -> Done / 已生成: {os.path.basename(new_path)}")

                    elif file_ext == ".xls":
                        # Lazy-load Excel / 懒加载 Excel
                        if excel is None:
                            excel = win32com.client.Dispatch("Excel.Application")
                            excel.Visible = False

                        wb = excel.Workbooks.Open(file_path)
                        new_path = os.path.splitext(file_path)[0] + ".xlsx"
                        # 51 = xlOpenXMLWorkbook (.xlsx)
                        wb.SaveAs(new_path, FileFormat=51)
                        wb.Close()
                        print(f"  -> Done / 已生成: {os.path.basename(new_path)}")

                except Exception as e:
                    print(f"  [Conversion failed / 转换失败]: {e}")

    # Cleanup resources / 清理资源
    if word: word.Quit()
    if excel: excel.Quit()

    print("\nUpgrade complete! Please re-run the main conversion script.")
    print("格式升级完成！请重新运行主转换脚本。")
    input("Press Enter to exit / 按回车退出...")

if __name__ == "__main__":
    target_dir = sys.argv[1] if len(sys.argv) > 1 else None
    convert_old_formats(target_dir)
import os
import sys
import win32com.client


def convert_old_formats(base_dir=None):
    """将目录中所有 .doc/.xls 文件批量升级为 .docx/.xlsx

    用法：python fix_old_files.py [目录路径]
    未指定目录时，默认使用脚本所在目录。
    """
    if base_dir is None:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    # 定义需要处理的旧格式后缀
    old_formats = {
        ".doc": ".docx",
        ".xls": ".xlsx"
    }
    
    print("正在启动 Office 组件，请耐心等待...")
    
    word = None
    excel = None
    
    # 遍历文件夹
    for root, dirs, files in os.walk(base_dir):
        for filename in files:
            if is_temporary_office_file(filename):
                continue
            file_ext = os.path.splitext(filename)[1].lower()
            file_path = os.path.join(root, filename)
            
            if file_ext in old_formats:
                print(f"发现旧格式文件: {filename}，正在转换...")
                
                try:
                    if file_ext == ".doc":
                        # 懒加载 Word
                        if word is None:
                            word = win32com.client.Dispatch("Word.Application")
                            word.Visible = False
                        
                        doc = word.Documents.Open(file_path)
                        # 新文件路径
                        new_path = os.path.splitext(file_path)[0] + ".docx"
                        # 16 代表 wdFormatXMLDocument (.docx)
                        doc.SaveAs2(new_path, FileFormat=16)
                        doc.Close()
                        print(f"  -> 已生成: {os.path.basename(new_path)}")
                        
                    elif file_ext == ".xls":
                        # 懒加载 Excel
                        if excel is None:
                            excel = win32com.client.Dispatch("Excel.Application")
                            excel.Visible = False
                            
                        wb = excel.Workbooks.Open(file_path)
                        new_path = os.path.splitext(file_path)[0] + ".xlsx"
                        # 51 代表 xlOpenXMLWorkbook (.xlsx)
                        wb.SaveAs(new_path, FileFormat=51)
                        wb.Close()
                        print(f"  -> 已生成: {os.path.basename(new_path)}")
                        
                except Exception as e:
                    print(f"  [转换失败]: {e}")

    # 清理资源
    if word: word.Quit()
    if excel: excel.Quit()
    
    print("\n格式升级完成！请重新运行主转换脚本。")
    input("按回车退出...")

if __name__ == "__main__":
    target_dir = sys.argv[1] if len(sys.argv) > 1 else None
    convert_old_formats(target_dir)
