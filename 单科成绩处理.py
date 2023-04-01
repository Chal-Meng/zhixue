from io import BytesIO
from typing import Dict, List
import pywebio
import pandas as pd

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
            pywebio.output.download(filename,io.getvalue())
        pywebio.output.put_tabs([{'title':filename,
                                'content':pywebio.output.put_html(df.to_html())}
                                for filename,df in data.items()])
    pywebio.output.put_button('开始处理',onclick=process)


pywebio.start_server(subject,port=12345,host='127.0.0.1')
