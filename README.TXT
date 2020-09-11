PDF2PPT使用说明
环境：windows10 + python3.6.5

0.安装依赖
pip install pdf2image
pip install python-pptx
到http://blog.alivate.com.au/wp-content/uploads/2018/10/poppler-0.68.0_x86.7z
下载poppler，并将bin加入环境变量

1.事先准备
<1>在PDF2PPT.py同目录下创建convert目录，并放入所有待转换的pdf文件到该文件夹内，不可放入其他类型文件
<2>在PDF2PPT.py同目录下新建幻灯片，命名为'template.pptx'，并打开并创建一张空白幻灯片，将模板幻灯片页面调整至需要的大小。
<3>如果需要封面，请将封面图片放到代码同目录下，并重命名为'0.jpg'。

2.运行程序Normal.py（无需封面）或Cover.py（需要封面）
生成输出文件夹'output'，所有PPT将以PDF的名称命名，并放入output/下
