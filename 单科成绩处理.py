print("""
#####成绩处理工具开始运行#####
请勿关闭这个窗口，否则不能完成工作。""")
from io import BytesIO
from typing import Dict
from functools import partial
import pywebio
import pandas as pd
from pyecharts.charts import Line
from pyecharts.options import ToolboxOpts,TitleOpts,LineItem,DataZoomOpts,AxisOpts,MarkLineItem,MarkLineOpts,LabelOpts,TooltipOpts
import webview

DELETE_COLUMNS=['序号','班级','自定义考号','准考证号','客观分','主观分']

@pywebio.config(title='成绩处理')
def main():
    #宽屏
    pywebio.session.set_env(output_max_width='95%')
    pywebio.output.put_buttons(['处理单科成绩','总分画图'],[process,total_plot])

def download(data:Dict[str,pd.DataFrame]):
    '''单科成绩下载'''
    for name,df in data.items():
        #检查扩展名
        if name.endswith('.xls'):name=name+'x'
        io=BytesIO()
        with pd.ExcelWriter(io,engine='xlsxwriter') as writer:
            df.to_excel(writer)
        io.seek(0)
        pywebio.output.download("（新）"+name,io.getvalue())
def process():
    """单科成绩处理"""
    files=pywebio.input.file_upload('上传成绩文件',accept='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel',placeholder='选择单科成绩文件。可以多选！',multiple=True)
    data:Dict[str,pd.DataFrame]={}

    #进度条
    pywebio.output.put_processbar('process',label='处理中')

    for file in files:
        filename=file['filename']
        if filename.endswith('.xls'):filename=filename+'x'#处理扩展名

        df=pd.read_excel(file['content'],header=1,index_col=1)
        df.drop(DELETE_COLUMNS,inplace=True,axis=1,errors='ignore')
        def highlight_max(s):
            try:is_max = s == s.max()
            except TypeError:return ['' for _ in range(len(s))]
            else:return ['font-style: italic; background-color: #bbbbbb' if v else '' for v in is_max]

        # 应用样式
        df = df.style.apply(highlight_max)
        data[filename] = df
        
        #进度条
        pywebio.output.set_processbar('process',1/len(files)*len(data))
    pywebio.output.put_tabs([
        {'title':filename,
        'content':[pywebio.output.put_html(df.to_html()),
                    plot_line_nian_ban(df.data)]
            }
                            for filename,df in data.items()])
    pywebio.output.put_button('下载',partial(download,data=data),color='success')

def total_plot():
    """总成绩画图"""
    files=pywebio.input.file_upload('上传总分成绩文件',accept='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel',placeholder='选择总成绩文件',multiple=True,help_text='可以多选，多选将生成时间轴')
    #data:Dict[str,pd.DataFrame]={}
    for file in files:
        plot_line_nian_ban(pd.read_excel(file['content']))

def plot_line_nian_ban(df:pd.DataFrame):
    """名次画图"""
    line=Line()
    
    #横轴
    ban_ci=df['班次'].to_list()
    line.add_xaxis(ban_ci)
    
    #纵轴
    data_show=[]
    for _,line_series in df.iterrows():
        data_show.append(LineItem(line_series.name,value=float(line_series['校次']),
                                label_opts=LabelOpts(rotate=-30,formatter='{b}',position='insideTopLeft'),
                                tooltip_opts=TooltipOpts(formatter='{b}：第{c}名，%d分'%line_series['总分'])))
    #标记线
    
    line.add_yaxis('成绩',data_show)
    
    #配置
    line.set_global_opts(toolbox_opts=ToolboxOpts(True),
                        title_opts=TitleOpts(title='名次分析图',subtitle='年级名次和班级名次之间的关系'),
                        datazoom_opts=DataZoomOpts(range_start=0,range_end=100),
                        xaxis_opts=AxisOpts(name='班级排名'),
                        yaxis_opts=AxisOpts(name='年级排名'),)
    return pywebio.output.put_html(line.render_notebook())


from threading import Thread

import asyncio
def start_thread():
    loop=asyncio.get_event_loop_policy().new_event_loop()
    asyncio.set_event_loop(loop)
    pywebio.start_server(main,port=12345,host='127.0.0.1')

Thread(target=start_thread).start()
webview.create_window("成绩处理工具", "http://127.0.0.1:12345",zoomable=True)
webview.start()

from os import _exit
_exit(0)
