#################################################################
# 指定されたフォルダ配下のExcelを開いていきシートのズーム倍率を揃えます.
#
# 実行には、以下のライブラリが必要です.
#   - win32com
#     - $ python -m pip install pywin32
#
# [参考にした情報]
#   - https://stackoverflow.com/a/37635373
#################################################################


# noinspection SpellCheckingInspection
def go(target_dir: str, zoom: int):
    import pathlib

    import pywintypes
    import win32com.client

    excel_dir = pathlib.Path(target_dir)
    if not excel_dir.exists():
        print(f'target directory not found [{target_dir}]')
        return

    if zoom <= 0:
        print(f'illegal zoom value [{zoom}')
        return

    try:
        excel = win32com.client.Dispatch('Excel.Application')
        excel.Visible = True

        for f in excel_dir.glob('**/*.xlsx'):
            abs_path = str(f)
            try:
                wb = excel.Workbooks.Open(abs_path)
            except pywintypes.com_error as err:
                print(err)
                continue

            try:
                sheets_count = wb.Sheets.Count
                for sheet_index in range(0, sheets_count):
                    ws = wb.Worksheets(sheet_index + 1)
                    ws.Activate()
                    excel.ActiveWindow.Zoom = zoom
                wb.Save()
                wb.Saved = True
            finally:
                wb.Close()
    finally:
        excel.Quit()


if __name__ == '__main__':
    target_directory = r'/path/to/excel/dir'
    zoom_value = 70

    go(target_directory, zoom_value)
