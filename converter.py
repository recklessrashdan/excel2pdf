import os
import win32com.client
from datetime import datetime

excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False

folder = os.getcwd()

log_file = "conversion_log.txt"

with open(log_file, "a", encoding="utf-8") as log:
    log.write("\n===== New Session =====\n")
    log.write(f"Time: {datetime.now()}\n\n")

for file in os.listdir(folder):
    if file.lower().endswith((".xlsx", ".xls")):

        input_path = os.path.join(folder, file)
        output_file = os.path.splitext(file)[0] + ".pdf"
        output_path = os.path.join(folder, output_file)

        try:
            wb = excel.Workbooks.Open(input_path)
            wb.ExportAsFixedFormat(0, output_path)
            wb.Close(False)

            print(f"Converted: {file}")

            with open(log_file, "a", encoding="utf-8") as log:
                log.write(f"SUCCESS: {file} -> {output_file}\n")

        except Exception as e:
            print(f"Failed: {file}")

            with open(log_file, "a", encoding="utf-8") as log:
                log.write(f"FAILED: {file} -> {e}\n")

excel.Quit()

print("\nAll done! Check conversion_log.txt")