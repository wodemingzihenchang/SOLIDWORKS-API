---
title: SwDocumentMgr注册
tag: SOLIDWORKS
toc: true
---

# SwDocumentMgr注册

新电脑使用时的一些问题和设置：手动注册卸载 SwDocumentMgr.dll：

## 库注册

SwDocumentMgr.dll 和 zlib.dll 必须位于同一文件夹中，或者可以通过 Windows 搜索路径访问。

单击开始，运行。在对话框中，键入：注册

```
regsvr32  "<disk>：\Program Files\Common Files\SOLIDWORKS Shared\SwDocumentMgr.dll”
```

单击“确定”。再次单击“确定”。

注意：安装程序（如 InstallShield）可以自动为您注册 DLL。

## 库卸载

单击开始，运行。在对话框中，键入：

```
regsvr32 /u ”<disk>：\Program Files\Common Files\SOLIDWORKS Shared\SwDocumentMgr.dll”
```

单击“确定”。再次单击“确定”。

您可以重新分发 swDocumentMgr.dll 和 zlib.dll。

多版本SW时，实例化COM组件的选择

#  Help手册：

[SolidWorks.Interop.swdocumentmgr Namespace](https://help.solidworks.com/2022/english/api/swdocmgrapi/SolidWorks.Interop.swdocumentmgr~SolidWorks.Interop.swdocumentmgr_namespace.html)

[Key申请](https://www.solidworks.com/support/subscription/key-request/ )

