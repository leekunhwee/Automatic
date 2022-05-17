# <center>宜春“翼起防疫”二维码自动生成软件使用说明</center>

<center> 2022.05.16</center>

<center> 李健辉   南昌市经开区樵舍镇永强村  18710886343</center>

<br>

## 目录:

-  [写在最前](#写在最前)
-  [前期准备](#前期准备)
-  [开始运行](#开始运行)
-  [开发者须知](#开发者须知)

<br>

## 写在最前

本软件由李健辉和电子科技大学计算机专业大三学生王正仁共同开发，发行版软件请见本repository的release。

## 前期准备

<ol>

<li>宜春“翼起防控”系统需要九项信息，如下图。所以需要提前将受检人员的身份信息按照模板整理好。
受检人员的信息整理格式如下（“证件号码”的单元格式要设为文本，“居住区域”必须要按照小程序中提供的地址信息填写，不得有误！）：


<div align = "center">
<img src = ".\info\table.png"  width=60% alt = "受检人员信息表模板" title = "受检人员信息表模板">
</div>
<p align = "center"><b>受检人员信息表模板</b></p>

<li>先通过右击屏幕-显示设置，将电脑的分辨率调整到1920×1080，缩放改为125%；如果没有该分辨率，请选择1280×720，缩放为100%（请务必提前设置，否则软件将无法正常运行），再通过微信打开“翼起防控”小程序，然后打开如下图所示注册页面，并将以下页面截图（注意：要截取自己电脑显示的图，截图方法请看下一页）。

<div align = "center">
<img src = ".\info\pic.png"  width=30% alt = "注册页面截图" title = "注册页面截图">
</div>
<p align = "center"><b>注册页面截图</b></p>

</ol>

截图可以使用免费实用的截图软件PicPick：

<div align = "center">
<img src = ".\info\picpick.png"  width=60% alt = "免费实用的截图软件" title = "免费实用的截图软件">
</div>
<p align = "center"><b>免费实用的截图软件</b></p>

Picpick 官网: [下载地址](https://picpick.app/zh/download)

也可以通过链接找到安装包： [下载地址](http://gofile.me/6Ugsb/dVrJ8aGOY)

截图时，主要需要用到的快捷键是 ：`Alt + PrtScr`

这个快捷键可截取屏幕上的活动窗口。

其中：`PrtScr` 也就是 `Print Screen`，即键盘上的截屏键。

截图之后，将其保存到一个比较好找的路径下，例如软件所在的文件夹。

## 开始运行

解压文件夹，双击打开文件夹中的 `QRcodeSimple_Yiqi_1920x1080.exe` 或者 `QRcodeSimple_Yiqi_1280x720.exe` （依据自己电脑的分辨率选择），稍等片刻，即会出现欢迎界面，按回车可看到一些注意事项。

<div align = "center">
<img src = ".\info\running.png"  width=60% alt = "软件运行界面" title = "软件运行界面">
</div>
<p align = "center"><b>软件运行界面</b></p>

两次回车之后，会有一个窗口界面弹出，点击“选择Excel表”，然后选择前期准备好的受检人员信息表。

<div align = "center">
<img src = ".\info\select.png"  width=60% alt = "受检人员信息表选择界面" title = "受检人员信息表选择界面">
</div>
<p align = "center"><b>受检人员信息表选择界面</b></p>

选择完受检人员信息表后，再选择刚刚截取的截图，在保证注册页面不被遮挡的前提下，点击“确定”。

如果前面的操作没有问题，而且屏幕显示恰好满足要求，那么程序就会进入自动循环，并生成二维码图片，图片将以“姓名_身份证”的形式命名，所以不会出现重名的情况。

最后将在应用程序所在的文件夹下新建一个 `qrcode` 文件夹，并即将所有图片存放在该文件内，二维码图片底部也包含姓名和身份证。 



可通过以下链接下载本软件：http://gofile.me/6Ugsb/dVrJ8aGOY



祝使用愉快！


### **祝使用愉快！**

## 开发者须知

**（非开发者无需了解以下内容）**

### 1.	安装python3.4以上版本(推荐3.8)，并配置环境变量。

[Python 教程](https://www.runoob.com/python3/python3-install.html)

可通过Anaconda安装Python，conda install python=3.8

为避免后期报错，注意将以下内容添加到系统环境变量的Path

<path>\Anaconda3

<path>\Anaconda3\scripts

<path>\Anaconda3\Library\bin

### 2.	安装依赖包

方法：在cmd中（win+R  输入cmd  回车）输入

`pip install openpyxl` 回车

`pip install pyperclip` 回车

`pip install xlrd` 回车

`pip install pyautogui==0.9.50` 回车

`pip install opencv-python` 回车

`pip install pillow==6.2.1` 回车

`pip install requests` 回车

`pip install qrcode` 回车


【国内下载慢，可以在install 后面加 -i https://pypi.tuna.tsinghua.edu.cn/simple 换成国内镜像】

该项目对应的 [Github 库](https://github.com/leekunhwee/Automatic)

作为编程业余爱好者，我非常欢迎大家一起学习交流。


