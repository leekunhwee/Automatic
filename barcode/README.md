# <center>条形码快速生成软件使用说明</center>
<center> 2022.05.03</center>

<center> 李健辉   南昌市经开区樵舍镇永强村  18710886343</center>

<br>

## 目录:

-  [写在最前](## 写在最前)
-  [前期准备](## 前期准备)
-  [开始运行](## 开始运行)
-  [开发者须知](## 对于开发者（非开发者无需了解以下内容）)

<br>

## 写在最前

本使用说明文档，以及代码本身写得很匆忙，而且本人非计算机专业，很多地方做的不够规范，常有低级错误，如有不尽如人意之处，欢迎批评指出。

## 前期准备

由于上饶海普洛斯系统需要四项信息，分别是：姓名、身份证、电话、住址（由于时间匆忙，这里没考虑除身份证以外的其他情形，请见谅）。所以需要提前将受检人员的身份信息整理好。

<div align = "center">
<img src = ".\info\interface.png"  width=60% alt = "G海普洛斯注册界面" title = "海普洛斯注册界面">
</div>
<p align = "center"><b>海普洛斯注册界面</b></p>

受检人员的信息整理格式如下（身份证的单元格式要设为文本，下图只是示例，录入信息时一定要用真实、准确的信息，当受检人信息表中存在身份证或姓名出错时，程序会被迫中断）：

<div align = "center">
<img src = ".\info\table.png"  width=60% alt = "受检人员信息表模板" title = "受检人员信息表模板">
</div>
<p align = "center"><b>受检人员信息表模板</b></p>

## 开始运行

解压文件夹，双击打开文件夹中的 BarcodeSimple.exe 应用程序，然后将该文件夹最小化，稍等片刻，即会出现欢迎界面，按回车可看到一些注意事项。

<div align = "center">
<img src = ".\info\running.png"  width=60% alt = "软件运行界面" title = "软件运行界面">
</div>
<p align = "center"><b>软件运行界面</b></p>

三次回车之后，会有一个窗口界面弹出，点击“选择Excel表”，然后选择前期准备好的受检人员信息表。

<div align = "center">
<img src = ".\info\select.png"  width=60% alt = "受检人员信息表选择界面" title = "受检人员信息表选择界面">
</div>
<p align = "center"><b>受检人员信息表选择界面</b></p>

选择完受检人员信息表后，点击“确定”（注意要将注册页面停留在“添加受检人”那一页，且不能被遮挡）。之后，就会进入自动循环注册阶段，循环过程中，如果程序没有报错，或者没有出现“未找到匹配图片”时，请不要再进行任何其他操作，耐心等待程序运行结束。

<b>本软件的设计思路是：通过图片识别完成鼠标的定位，并给鼠标添加操作。经过多方测试发现，有的人虽然都按照步骤正确操作了，但在输入Excel表之后会就一直报“未找到匹配图片,0.1秒后重试，关闭程序按 Alt+F4 ”。这说明，文件夹在提供的截图与你用的电脑显示效果不匹配。为解决这一问题，请根据软件所在文件夹中的截图内容，将每一个截图都替换为自己的截图，注意要保存为 .png 格式（由于识别定位位置在图片的正中心，所以截图时请注意截图的正中心才是鼠标点击的位置，另外，10.png 只有在右击条形码的时候才会出现）。之后再重新打开文件夹中的BarcodeSimple.exe软件，看看是否能够匹配图片。</b>

条形码截图将以“姓名_身份证”的形式命名，所以不会出现重名的情况。默认是放在了“我的电脑（此电脑）-文档”文件夹中（根据每个人的情况会有不同）。

祝使用愉快！

## 对于开发者（非开发者无需了解以下内容）

1.	安装python3.4以上版本，并配置环境变量
[Python 教程](https://www.runoob.com/python3/python3-install.html)
2.	安装依赖包
方法：在cmd中（win+R  输入cmd  回车）输入
pip install openpyxl 回车
pip install pyperclip 回车
pip install xlrd 回车
pip install pyautogui==0.9.50 回车
pip install opencv-python 回车
pip install pillow 回车
【国内下载慢，可以在install 后面加 -i https://pypi.tuna.tsinghua.edu.cn/simple 换成国内镜像】

该项目对应的 [Github 库](https://github.com/leekunhwee/Automatic)

作为编程业余爱好者，我非常欢迎大家一起学习交流。


