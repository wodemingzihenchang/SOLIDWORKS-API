---
title: SOLIDWORKS-API后台打开文件
tag: SOLIDWORKS
toc: true
---

# SOLIDWORKS-API后台打开文件

## 关键方法

关闭文件显示

```C#
//关闭文件显示
swApp.DocumentVisible(false, (int)swDocumentTypes_e.swDocDRAWING);

//打开文件
ModelDoc2 drawing_com = swApp.OpenDoc6(drawing_path, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", errors, warnings);
......
    
//恢复文件显示
swApp.DocumentVisible(true, (int)swDocumentTypes_e.swDocDRAWING);
```



## 方法2

关闭文件显示

```C#
ModelDoc2.Visible = false;
```

