import os

import pandas as pd

# 設定
filename = "/Users/hirohito/Documents/work/repository/yorozu/20250202_split_file/work/知財①2024_0913_0927_22696_S8_9195.xlsm"
sheetname = "ダウンロード項目"
header = 2
filename_out_base = "知財①2024"
output_file_num = 1

df = pd.read_excel(filename, sheet_name=sheetname, header=header)
df = df.dropna(how="all", axis=1)

output_dir = "work/output_files"
os.makedirs(output_dir, exist_ok=True)

rows_per_split = len(df) // output_file_num
# 分割してjsonファイルで出力
for i in range(output_file_num):
    start_row = i * rows_per_split
    end_row = (i + 1) * rows_per_split if i != output_file_num - 1 else len(df)
    split_df = df.iloc[start_row:end_row]
    split_df.to_json(
        f"{output_dir}/{filename_out_base}_{i + 1}.txt",
        orient="records",
        force_ascii=False,
        indent=4,
    )
