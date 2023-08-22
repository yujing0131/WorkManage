import plotly.express as px
import plotly.graph_objects as go

from IPython.display import display, update_display, HTML
df_data = px.data.iris()
#print(df_data)
#直方圖
fig = px.histogram(df_data, x="sepal_width")
HTML(fig.to_html())
fig.write_html("./demo1.html")
fig = px.histogram(df_data, x="sepal_width", color="species")
fig.update_layout(barmode='overlay')
fig.update_traces(opacity=0.75)
HTML(fig.to_html())

##特徵關聯度分析
fig = px.scatter_matrix(df_data, dimensions=["sepal_width", "sepal_length", "petal_width", "petal_length"], color="species")
HTML(fig.to_html())
##散佈圖
fig = px.scatter(df_data, x="sepal_width", y="sepal_length")
HTML(fig.to_html())

##3維視覺化
fig = px.scatter_3d(df_data, x="sepal_width", y="sepal_length", z="petal_width", 
                    color="species",size='petal_length')
HTML(fig.to_html())

##參考資料:https://ithelp.ithome.com.tw/articles/10277258