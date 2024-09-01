import pandas as pd
from datetime import datetime
import plotly.graph_objects as go
# import plotly.express as px


path = 'C:\\Users\\152551\\Desktop\\report_2022Q1_raredrug.csv'
df = pd.read_csv(path, encoding='utf-8')

df1 = df[['Brand_name', 'cost_private']]
df1 = df1.sort_values(by='cost_private', ascending=True)

# fig = px.bar(df1, y='Brand_name', x='cost_private', text='cost_private')
# fig.update_traces(texttemplate='%{text:.2s}', textposition='outside')
# fig.update_layout(uniformtext_minsize=8, uniformtext_mode='hide')
# fig.show()

titlename = '臨採藥品金額報表 ' + '2021年 11月1日 至 2022年 1月31日'

fig = go.Figure(go.Bar(
    x=df1['cost_private'],
    y=df1['Brand_name'],
    orientation='h',
))

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

total_cost = df1['cost_private'].sum()

drug = df1['Brand_name']
money_drug = df1["cost_private"].astype('int')

annotations1 = []

for yd, xd in list(zip(money_drug, drug)):
    # labeling the money
    annotations1.append(dict(xref='x1', yref='y1',
                             # y=xd, x=yd + 1,
                             y=xd, x=yd + 12,  # x=yd + 2 處方籤數量註記與長條圖的距離為 2
                             text="$ " + f'{yd:,}',
                             # font=dict(family='Arial', size=14, color='rgb(65,105,225)'), showarrow=False))
                             font=dict(family='consolas', size=15,
                                       color='rgb(0,0,205)'),
                             showarrow=False,
                             xshift=45,
                             align="right"))

# labeling the comment
annotations1.append(dict(xref='paper', yref='paper',
                         x=0.01, y=0.96,
                         # x=-0.2, y=0.109,
                         text="總購藥費用: $" + f'{total_cost:,}',
                         font=dict(family='consolas', size=16,
                                   color='rgb(0,0,205)'),
                         showarrow=False))

fig.update_layout(annotations=annotations1,
                  margin=dict(l=100, r=50, b=50, t=50, pad=4))
# fig.update_annotations(align="right")

now = datetime.now()
quarter_of_the_year = 'Q' + str((now.month - 1) // 3 + 1)

outputname = 'report_' + str(now.year) + quarter_of_the_year + '_raredrug'
html_path = 'C:\\Users\\152551\\Desktop\\' + outputname + '2.html'
img_path = 'C:\\Users\\152551\\Desktop\\' + outputname + '2.png'
print(html_path)
print(img_path)

fig.write_html(html_path)
fig.write_image(img_path)

msg2 = '作圖完成，共兩個檔案(csv, html)。' + '\n檔案名稱：' + outputname + '，存在桌面'
print(msg2)
