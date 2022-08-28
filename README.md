# SOLIDWORKS-API

## 介绍

Sw_toolkit是使用SOLIDWORKS API进行的二次开发工具，意在解决一些非设计性的操作，提高工程师的设计效率。
<img src="https://wodemingzihenchang.github.io/Project/index/SW API 窗口.png">

## 使用

1、下载程序：从Code按键下载ZIP压缩包，解压得到Sw_toolkit文件夹。

2、加载命令：打开SOLIDWORKS——》新建一个空白模板——》打开自定义功能——》找到“定义宏”命令——》拖放至空白工具栏——》选择Sw_toolkit文件夹RunMacro(progID).swp——》完成

3、测试效果：检查宏按钮能否打开Sw_toolkit工具。

## 功能

<a href="https://www.bilibili.com/video/BV1yd4y1m7sm/?spm_id_from=333.788&vd_source=61426b94de87e3e2baee96fc2f5b14f2">视频介绍</a>

<strong>输出属性：</strong>配合【选择文件】对多个零件的属性导出成Excel表。如果未选择文件则输出当前文件的属性

<strong>添加属性：</strong>从Excel表中获得属性并添加到零件（Excel表的格式有要求，请参考【输出属性】的Excel）

<strong>删除属性：</strong>配合【选择文件】对多个零件的属性进行删除。如果未选择文件则删除当前文件的属性

<strong>装配体分配多配置：</strong>对材料明细表进行配置分配，在多配置下可关联配置变化明细表

<strong>冻结零部件：</strong>对装配体零部件特征进行冻结，优化装配体的运行性能（最后是在修订完后进行该操作，设计中还是需解冻的）

<strong>需求更新中：</strong>就是需求更新中....

<strong>工程图输出多格式：</strong>配合【选择文件】对多个工程图，按需输出多种格式的文件。

<strong>工程图Bom输出XML：</strong>对材料明细表的信息转成.XML文件，文件保存在相同目录下。

<strong>工程图线条改颜色分层：</strong>对装配体工程图的顶层零部件进行分层并添加颜色区分。

<strong>批量工具，选择文件：</strong>就是选则文件。

<strong>批量工具，选择程序：</strong>就是选进行批量处理的SW宏命令程序。

<strong>批量工具，启动按钮：</strong>选中好零件和宏程序后进行启动批量处理。
