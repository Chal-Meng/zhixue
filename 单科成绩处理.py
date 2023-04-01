print("""
#####成绩处理工具开始运行#####

请在浏览器输入这个网址：http://127.0.0.1:12345/

在网页关闭之前，请勿关闭这个窗口，否则不能完成工作。""")
from io import BytesIO
from typing import Dict
import pywebio
import pandas as pd
from pyecharts.charts import Line
from pyecharts.options import ToolboxOpts,TitleOpts,LineItem

DELETE_COLUMNS=['Unnamed: 2']

@pywebio.config(title='单科成绩处理',description='单科成绩突出显示满分的；显示图片')
def subject():
    def download(data:Dict[str,pd.DataFrame]):
        for name,df in data.items():
            #检查扩展名
            if name.endswith('.xls'):name=name+'x'
            io=BytesIO()
            with pd.ExcelWriter(io,engine='xlsxwriter') as writer:
                df.to_excel(writer)
            io.seek(0)
            pywebio.output.download("（新）"+name,io.getvalue())
    def process():
        files=pywebio.input.file_upload('上传成绩文件',accept='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel',placeholder='选择单科成绩文件。可以多选！',multiple=True)
        data:Dict[str,pd.DataFrame]={}
        for file in files:
            filename='(新)'+file['filename']
            if filename.endswith('.xls'):filename=filename+'x'#处理扩展名

            df=pd.read_excel(file['content'])
            df.drop(DELETE_COLUMNS,inplace=True,axis=1,errors='ignore')
            def highlight_max(s):
                is_max = s == s.max()
                return ['font-style: italic; background-color: #bbbbbb' if v else '' for v in is_max]

            # 应用样式
            
            df = df.style.apply(highlight_max)
            data[filename] = df
        
            io=BytesIO()
            with pd.ExcelWriter(io,'xlsxwriter') as writer:
                df.to_excel(writer)
            io.seek(0)
            #pywebio.output.download(filename,io.getvalue())
        pywebio.output.put_tabs([{'title':filename,
                                'content':pywebio.output.put_html(df.to_html())}
                                for filename,df in data.items()])
        pywebio.output.put_button('下载',download,color='success')
    pywebio.output.put_button('开始处理',onclick=process)

@pywebio.config(title='总成绩画图',description='给总成绩画图。时间轴+排名分析+偏科对比')
def total_plot():
    files=pywebio.input.file_upload('上传成绩文件',accept='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel',placeholder='选择总成绩文件',multiple=True,help_text='可以多选，多选将生成时间轴')
    data:Dict[str,pd.DataFrame]={}
    for file in files:
        
        df=pd.read_excel(file['content'])
        df.drop(DELETE_COLUMNS,inplace=True,axis=1,errors='ignore')
        data[file['filename']] = df
        #名次画图
        line=Line()
        ban_ci=df['班次'].to_list()
        line.add_xaxis(ban_ci)
        data_show=[]
        for _,line_series in df.iterrows():
            data_show.append(LineItem(line_series['姓名'],line_series['校次']))
        line.add_yaxis('成绩',data_show)
        line.set_global_opts(toolbox_opts=ToolboxOpts(True),title_opts=TitleOpts(title='名次分析图'))
        pywebio.output.put_html(line.render_notebook())
        
        #成绩雷达图
        full_grades=pd.Series({'数学':120,'语文':120,'英语':120,'文综':120,'理综':120})
        grade_float=df/full_grades
        
pywebio.start_server([subject,total_plot],port=12345,host='127.0.0.1')
