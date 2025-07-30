import os
import pandas as pd

# 讀取 Excel 表格
base_dir = os.getcwd()  # ← 這裡改掉 __file__
file_name = '簡化名稱.xlsx'
file_path = os.path.join(base_dir, file_name)

# 讀取檔案
items_df = pd.read_excel(file_path)

# 輸入訂單資料檔案
file = input('請輸入訂單檔案路徑+檔名+類型: ')
df = pd.read_excel(file)

# 電話補 0
df['電話號碼'] = df['電話號碼'].astype(str).str.zfill(10)
df['收件人電話號碼'] = df['收件人電話號碼'].astype(str).str.zfill(10)

# 儲存整理後資料
output_rows = []

# 對每一筆訂單進行整理
for number in df['訂單號碼'].unique():
    flitered = df[df['訂單號碼'] == number]

    # 顧客基本資訊
    selected = flitered[['顧客','電話號碼','收件人', '收件人電話號碼','城市','地址 2','地址 1']].drop_duplicates().iloc[0]
    訂購資訊文字 = f"{selected['顧客']}  {selected['電話號碼']}  {selected['收件人']}  {selected['收件人電話號碼']}\n{selected['城市']}{selected['地址 2']}{selected['地址 1']}"

    # 預埋螺母資訊
    result = ''
    lowmu_texts = flitered['自訂訂單欄位 3 (是否要預埋螺母？)'].dropna().unique()
    for txt in lowmu_texts:
        if '⭕ 要預埋螺母' in str(txt):
            result = '⭕ 要預埋螺母'
            break

    # 商品簡化名稱處理（有空白則用商品名稱）
    names_list = []
    for _, row in flitered.iterrows():
        item_code = row['商品貨號']
        item_name = row['商品名稱']
        if pd.isna(item_code) or str(item_code).strip() == '':
            names_list.append(str(item_name))
        elif item_code.upper() == 'CD01':   # ← 新增：貨號為 CD01 時統一用商品名稱
            names_list.append(item_name)
        else:
            match = items_df[items_df['代號'] == item_code]
            if not match.empty and pd.notna(match.iloc[0]['簡化名稱']):
                names_list.append(str(match.iloc[0]['簡化名稱']))
            else:
                names_list.append(str(item_name))
                
    number_list = []
    quantities = flitered['數量'].tolist()
    for qty in quantities:
        if qty > 1:
            number_list.append(f"*{qty}")
        else:
            number_list.append('')
        
    combined_list = []
    for name, qty in zip(names_list, number_list):
        combined_list.append(f"{name}{qty}")

    names = '+'.join(combined_list)

 
    # 付款與送貨資訊
    selected_price = flitered[['付款方式','付款總金額']].drop_duplicates().iloc[0]
    selected_way = flitered[['送貨方式']].drop_duplicates().iloc[0]
    note1 = flitered['自訂訂單欄位 1 (店主備註)'].dropna().unique()
    note2 = flitered['自訂訂單欄位 2 (🔺配件位置)'].dropna().unique()

    # 合併成一列
    row = {
        '訂單號碼': number,
        '送貨方式': selected_way['送貨方式'],
        '訂購資訊': 訂購資訊文字 + ('\n' + result if result else ''),
        '品項': names,
        '件數':'',
        '付款方式': selected_price['付款方式'],
        '付款總金額': selected_price['付款總金額'],
        '出圖':'',
        '店主備註': note1[0] if len(note1) > 0 else '',
        '配件位置': note2[0] if len(note2) > 0 else '',
    }

    output_rows.append(row)

# 轉成 DataFrame 並輸出成 Excel
output_df = pd.DataFrame(output_rows)
output_df.to_excel('output.xlsx', index=False)
print("✅ 已成功匯出成 output.xlsx！")
