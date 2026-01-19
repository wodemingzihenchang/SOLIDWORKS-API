---
title: Sw_模型尺寸标注
tag: SOLIDWORKS
toc: true
---

# Sw_模型尺寸标注

模型尺寸标注

<!-- more -->

## 操作

使用【注解-模型项目】倒是可以实现自动获得模型草图的尺寸标注。

## 代码

```c#
//新建图纸文件
string Templatepath = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateDrawing);
swModel = (ModelDoc2)swApp.NewDocument(Templatepath, 0, 0, 0);
//从文件路径创建标准3视图
swDrawing.Create3rdAngleViews(filepath);
//插入模型项目
swDrawing.InsertModelAnnotations3(0, 32768, true, false, false, false);
```

## 使用

参考：作者：Paine Zeng API-尺寸链标注链接：https://zhuanlan.zhihu.com/p/597967278

````vb
Dim swModel As ModelDoc2
Dim swDrawing As DrawingDoc
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long

Sub main()

Set swApp = Application.SldWorks
Set swModel = swApp.ActiveDoc

Set swDrawing = swApp.ActiveDoc
Dim vAnnotations As Variant
vAnnotations = swDrawing.InsertModelAnnotations3(0, 32768, True, False, False, False)

swModel.Save

End Sub
````

