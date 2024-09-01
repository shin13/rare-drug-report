import pandas as pd
from datetime import datetime
import plotly.graph_objects as go

# (1) 更新臨採清單 C:\Users\152551\Documents\Raredrug\Raredrug_ATC_Ref_final.xlsx
# (2) 下載批價扣庫明細，用csv with headers（編碼UTF8）格式儲存
# (3) 更新 C:\Users\152551\Documents\Raw\1.藥品基本檔.xlsx
# (4) 執行程式 report.py
# (5) 完成的檔案在桌面
# (6) 確認圖片X軸長度是否需要調整，如須調整則修改後重製

print("請先更新(1)臨採清單、(2)藥品基本檔，並下載批價扣庫明細。")
path = input("請將批價扣庫明細的csv with headers檔案（編碼UTF8）拖曳至視窗中後按ENTER。")
# path = "C:\\Users\\152551\\Downloads\\202206-08_stock.csv"
# 請下載成CSV with headers
print("檔案處理中...")
# df = pd.read_csv(path, encoding='big5hkscs')
df = pd.read_csv(path, encoding="utf-8")

df1 = df.loc[df["trans_id"] == "0"]
df1 = df1[["patient_id", "visit_date", "code_id", "tot_qty"]]
df1["visit_date"] = pd.to_datetime(df1["visit_date"].astype(str), format="%Y%m%d")
df1["code_id"] = df1["code_id"].str.strip()
df1.rename(columns={"code_id": "drug_code"}, inplace=True)
df1["pt_date"] = df1["patient_id"].astype(str) + "_" + df1["visit_date"].astype(str)
df_agg = (
    df1.groupby("drug_code")
    .agg(qty=("tot_qty", sum), count=("pt_date", "nunique"))
    .reset_index(drop=False)
)

# 讀取臨採藥品清單
ref = pd.read_excel("Raredrug_ATC_Ref_final.xlsx")
ref["drug_code"] = ref["drug_code"].str.strip()
merge = pd.merge(df_agg, ref, how="inner", on="drug_code")


# basic_path = r'C:\Users\152551\Documents\Raw\drug_basic.csv'
basic_path = "C:\\Users\\152551\\Documents\\Raw\\1.藥品基本檔.xlsx"
print(basic_path)
# basic = pd.read_csv(basic_path, encoding='big5hkscs')
basic = pd.read_excel(basic_path)
price_ref = basic[["藥品代碼", "自費價格", "健保價格"]]
price_ref.rename(
    columns={"藥品代碼": "drug_code", "自費價格": "price_private", "健保價格": "price_nhi"},
    inplace=True,
)

# 把用量跟金額融合
output = pd.merge(merge, price_ref, how="left", on="drug_code")
output["price_private"] = output["price_private"].astype(float)
output["cost_private"] = output["qty"] * output["price_private"]
output["cost_nhi"] = output["qty"] * output["price_nhi"]
output = output[
    [
        "drug_code",
        "qty",
        "count",
        "Brand_name",
        "compound_name",
        "price_private",
        "price_nhi",
        "cost_private",
        "cost_nhi",
    ]
]
output.sort_values(by=["count", "cost_private"], ascending=False, inplace=True)

# 計算各個藥品的金額百分比
total_cost = output["cost_private"].sum().astype("int")
output["cost_pct"] = output["cost_private"] / total_cost
output["cost_pct"] = output["cost_pct"].apply(lambda x: f"{x:.2%}")
output["cost_private2"] = output["cost_private"].apply(lambda x: f"{x:,.0f}")
output["cost_anno"] = (
    output["cost_private2"].astype("str") + " (" + output["cost_pct"] + ")"
)

now = datetime.now()
quarter_of_the_year = "Q" + str((now.month - 1) // 3 + 1)
filename = "report_" + str(now.year) + quarter_of_the_year + "_raredrug.csv"
filepath = "C:\\Users\\152551\\Desktop\\" + filename
output.to_csv(filepath, index=False)

msg = "檔案處理完成，作圖中..."
print(msg)

title = "臨採藥用量與金額報表 "
if quarter_of_the_year == "Q1":
    titlename = (
        title + str(now.year - 1) + "年" + " 12月1日 至 " + str(now.year) + "年" + " 2月28日"
    )
elif quarter_of_the_year == "Q2":
    titlename = (
        title + str(now.year) + "年" + " 3月1日 至 " + str(now.year) + "年" + " 5月31日"
    )
elif quarter_of_the_year == "Q3":
    titlename = (
        title + str(now.year) + "年" + " 6月1日 至 " + str(now.year) + "年" + " 8月31日"
    )
elif quarter_of_the_year == "Q4":
    titlename = (
        title + str(now.year) + "年" + " 9月1日 至 " + str(now.year) + "年" + " 11月30日"
    )
else:
    print("季別計算結果為：" + quarter_of_the_year)

df_plot = pd.read_csv(filepath)
# 因為IPEY Peyona第三季是負的，用3支，回帳-6，所以先改成正值。
# df_plot.loc[29, ['qty', 'cost_private']] = [3, 3060.0]
df_plot = df_plot[["Brand_name", "count", "cost_private", "cost_anno"]]


fig = go.Figure(
    go.Bar(
        x=df_plot["count"],
        y=df_plot["Brand_name"],
        orientation="h",
    )
)

fig.update_layout(
    title=titlename,
    yaxis=dict(
        showgrid=False,
        showline=False,
        showticklabels=True,
        domain=[0, 0.9],
    ),
    xaxis=dict(
        zeroline=False,
        showline=False,
        showticklabels=True,
        showgrid=True,
        domain=[0, 0.8],
    ),
    # font=dict(size=15),
    margin=dict(l=50, r=50, b=50, t=50, pad=4),
)

total_count = df_plot["count"].sum()
total_cost = df_plot["cost_private"].sum()

drug = df_plot["Brand_name"]
# df_plot["money_annotation"] = "$" + df_plot["cost_private"].astype('str')
# money_drug = df_plot["money_annotation"]
df_plot["cost_private"] = df_plot["cost_private"].fillna(0)
money_drug = df_plot["cost_anno"].astype("str")
count_drug = df_plot["count"]


annotations1 = []

for ydn, yd, xd in list(zip(money_drug, count_drug, drug)):
    # labeling the money
    annotations1.append(
        dict(
            xref="x1",
            yref="y1",
            # y=xd, x=60,  # 使用x調整annotaion的位置 y則比照xd(Brand_name的位置)
            # 使用x調整annotaion的位置 y則比照xd(Brand_name的位置)
            y=xd,
            x=75,
            # text=str(ydn),
            text="$ " + f"{ydn}",
            font=dict(family="roboto", size=15, color="rgb(0,0,205)"),
            showarrow=False,
            align="right",
        )
    )
    # labeling the count
    annotations1.append(
        dict(
            xref="x1",
            yref="y1",
            # y=xd, x=yd + 1,
            y=xd,
            x=yd + 3,  # x=yd + 2 處方籤數量註記與長條圖的距離為 2
            text=str(yd),
            # font=dict(family='Arial', size=14, color='rgb(65,105,225)'), showarrow=False))
            font=dict(family="roboto", size=15, color="rgb(0,0,205)"),
            showarrow=False,
            align="right",
        )
    )

# labeling the comment
annotations1.append(
    dict(
        xref="paper",
        yref="paper",
        x=0.01,
        y=0.96,
        # x=-0.2, y=0.109,
        text="總處方張數: "
        + f"{total_count:,.0f}"
        + "      "
        + "總購藥費用: $"
        + f"{total_cost:,.0f}",
        font=dict(family="roboto", size=16, color="rgb(0,0,205)"),
        showarrow=False,
    )
)

fig.update_layout(annotations=annotations1, margin=dict(l=100, r=50, b=50, t=50, pad=4))
# fig.update_annotations(align="right")

outputname = "report_" + str(now.year) + quarter_of_the_year + "_raredrug"
html_path = "C:\\Users\\152551\\Desktop\\" + outputname + ".html"
img_path = "C:\\Users\\152551\\Desktop\\" + outputname + ".png"
print(html_path)
print(img_path)

fig.write_html(html_path)
fig.write_image(img_path)

msg2 = "作圖完成，共兩個檔案(csv, html)。" + "\n檔案名稱：" + outputname + "，存在桌面"
print(msg2)
