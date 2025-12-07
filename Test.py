import pandas as pd
input_file = "1205.xlsx"
output_file = "1.xlsx"
df = pd.read_excel(input_file, header=None, dtype=str)
TARGET_COLS = ["工號", "姓名", "部門", "部門名稱", "歸屬日期", "最早刷卡時間", "最晚刷卡時間", "時數"]

def is_header_row(row):
    return all(col in row.values for col in TARGET_COLS[:-1])
header_indices = [i for i, row in df.iterrows() if is_header_row(row)]
header = df.iloc[header_indices[0]].tolist()
data_rows = []
for start in header_indices:
    for _, row in df.iloc[start+1:].iterrows():
        if any("門禁刷卡查詢" in str(x) for x in row if pd.notna(x)):
            continue
        if any(str(x).startswith("備註說明") for x in row if pd.notna(x)):
            continue
        text = "".join(str(x) for x in row if pd.notna(x))
        if text.strip() == "":
            continue
        if is_header_row(row):
            break
        data_rows.append(row.tolist())
clean_df = pd.DataFrame(data_rows, columns=header)
clean_df = clean_df[TARGET_COLS]

clean_df.to_excel(output_file, index=False)
print("Done!")
