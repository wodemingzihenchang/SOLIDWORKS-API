---
title: Sw_导出钣金展开图
tag: SOLIDWORKS
toc: true
---

# Sw_导出钣金展开图

在solidworks做好的钣金零件，我们通常还需要进行需要将该钣金的展开图纸做导出2D图使用。那么就有了如下的方法：

## 使用

该工具用于批量导出钣金展开图，格式默认为DXF。目的是为了解决在处理数量多时的导出操作，有点类似SW自带的Task Schedule工具。

- **选择文件**

选择需要处理的钣金零件（非钣金无法导出DXF）。判断是否为钣金零件，可依据零件能否展开，特征树中是否有“钣金”特征文件夹。

<img src="https://i0.hdslb.com/bfs/article/aac60800fb478f113d6bda48444123a9025788d2.png@!web-article-pic.webp">

- **运行**

运行导出操作，这里是会操作到SW软件的。所以你可能需要停下你手头上的设计工作。去倒杯茶等待下软件自动完成。

<img src="https://i0.hdslb.com/bfs/article/2980f6877d012a0ff64de5214c56ca46bb710c65.png@1256w_666h_!web-article-pic.webp">

## 代码

```C# 
//获取当前已打开的零件
PartDoc swPart = (PartDoc)swDoc;
//获取保存路径
string swDocName = swDoc.GetPathName();
string swDxfName = swDocName.Substring(0, swDocName.Length - 6) + "dxf";
swPart.ExportToDWG2(swDxfName, swDocName, (int)swExportToDWG_e.swExportToDWG_ExportSheetMetal, true, null, false, false, 2149, null);
#region
```



## 平展设置

这里是钣金相关的设置

<img src="https://i0.hdslb.com/bfs/article/6f7754db551980777da31bb7d2dc1f1e495e11af.png@1256w_772h_!web-article-pic.webp">

你可以在“平板型式颜色”中提前设置好导出的图形的颜色，便于你阅读图纸。

<img src="https://i0.hdslb.com/bfs/article/eb73284a10d35ee4f25827493dcb996d756056be.png@!web-article-pic.webp">

**程序下载**

https://gitee.com/littleboy97/solidworks-api

