---
title: Sw_事件
tag: SOLIDWORKS
toc: true
---

# Sw_事件

在零件PartDoc，装配AssemblyDoc，图纸DrawingDoc类里可以使用事件进行监视。

以保存时进行内容检查为例：

```C#
//事件触发
pDoc.FileSaveNotify += Save_FileSaveNotify;
//事件触发后，进行的操作方法
private static int Save_FileSaveNotify(){}
```

## 代码

```C#
private static SldWorks swApp = (SldWorks)SwModelDoc.ConnectToSolidWorks();
private static PartDoc pDoc;
private static AssemblyDoc aDoc;
private static DrawingDoc dDoc;

//保存事件
public static void SaveEvent()
{
    ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc; 

    if (swModel.GetType() == (int)swDocumentTypes_e.swDocPART) { pDoc = (PartDoc)swModel; }
    else if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY) { aDoc = (AssemblyDoc)swModel; }
    else if (swModel.GetType() == (int)swDocumentTypes_e.swDocDRAWING) { dDoc = (DrawingDoc)swModel; }
	//监视事件
    if ((pDoc != null)) { pDoc.FileSaveNotify += Save_FileSaveNotify; }
    if ((aDoc != null)) { aDoc.FileSaveNotify += Save_FileSaveNotify; }
    if ((dDoc != null)) { dDoc.FileSaveNotify += Save_FileSaveNotify; }
}
//触发后的操作
private static int Save_FileSaveNotify()
{
    MessageBox.Show("触发保存事件");return 0;
}
```

## 使用

事件监控并不会凭空出现，事件方法需要通过加载到后台程序里（通过挂在后台的前提下，这样在对SW程序操作时就能捕获事件，来执行方法）。

<img src="README\使用.png">



# 其他

## 插件启动

此方法来自与[Creating C# add-in for SOLIDWORKS automation using API (codestack.net)](https://www.codestack.net/solidworks-api/getting-started/add-ins/csharp/)。

API接口来自与：ISwAddin 。其方法成员如下：

- ConnectToSW：加载外接程序时调用此方法。 
- DisconnectFromSW：当 SOLIDWORKS 即将被销毁时调用此方法。

1、注册表注册插件，

```C#
private const string ADDIN_KEY_TEMPLATE = @"SOFTWARE\SolidWorks\Addins\{{{0}}}";
private const string ADDIN_STARTUP_KEY_TEMPLATE = @"Software\SolidWorks\AddInsStartup\{{{0}}}";
private const string ADD_IN_TITLE_REG_KEY_NAME = "Title";
private const string ADD_IN_DESCRIPTION_REG_KEY_NAME = "Description";

[ComRegisterFunction]
public static void RegisterFunction(Type t)
{
    try
    {
        var addInTitle = "";
        var loadAtStartup = true;
        var addInDesc = "";
        var dispNameAtt = t.GetCustomAttributes(false).OfType<DisplayNameAttribute>().FirstOrDefault();

        if (dispNameAtt != null){addInTitle = dispNameAtt.DisplayName;}
        else{addInTitle = t.ToString();}

        var descAtt = t.GetCustomAttributes(false).OfType<DescriptionAttribute>().FirstOrDefault();

        if (descAtt != null){addInDesc = descAtt.Description;}
        else{addInDesc = t.ToString();}

        var addInkey = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(
   string.Format(ADDIN_KEY_TEMPLATE, t.GUID));

        addInkey.SetValue(null, 0);

        addInkey.SetValue(ADD_IN_TITLE_REG_KEY_NAME, addInTitle);
        addInkey.SetValue(ADD_IN_DESCRIPTION_REG_KEY_NAME, addInDesc);

        var addInStartupkey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(
   string.Format(ADDIN_STARTUP_KEY_TEMPLATE, t.GUID));

        addInStartupkey.SetValue(null, Convert.ToInt32(loadAtStartup), Microsoft.Win32.RegistryValueKind.DWord);
    }
    catch (Exception ex)
    {
        Console.WriteLine("Error while registering the addin: " + ex.Message);
    }
}

[ComUnregisterFunction]
public static void UnregisterFunction(Type t)
{
    try
    {
        Microsoft.Win32.Registry.LocalMachine.DeleteSubKey(
   string.Format(ADDIN_KEY_TEMPLATE, t.GUID));

        Microsoft.Win32.Registry.CurrentUser.DeleteSubKey(
   string.Format(ADDIN_STARTUP_KEY_TEMPLATE, t.GUID));
    }
    catch (Exception e)
    {
        Console.WriteLine("Error while unregistering the addin: " + e.Message);
    }
}
```



2、启动与关闭函数



```C#
private ISldWorks m_App;
//打开SW
//（指向 SldWorks Dispatch 对象的指针，加载项 ID）
public bool ConnectToSW(object ThisSW, int Cookie)
{
    m_App = ThisSW as ISldWorks;
    m_App.SendMsgToUser("Open SW");
    return true;
}
//关闭SW
public bool DisconnectFromSW()
{
    m_App.SendMsgToUser("Close SW");
    return true;
}
```



## 自启动宏

快捷方式的自启动宏方法：原本想法是通过打开SW程序时，自动加载事件监控后台。但这个方法还没得到验证，需要测试下。

问题1：局限性大，只有在用快捷方式打开时，才会有效果，然后双击打开的零件似乎就无效果。

```
"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\SLDWORKS.EXE" /m "C:\My Macros\Macro1.swp"
```



![Shortcut with macro path](https://www.codestack.net/solidworks-api/getting-started/macros/run-macro-on-solidworks-start/shortcut-with-macro-run.png)



 