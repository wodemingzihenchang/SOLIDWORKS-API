---
title: 创建属性页
tag: SOLIDWORKS
toc: true
---

# 创建属性页



<img src="README\属性页示例.png">

<!--more-->

## 代码实例：

演示一个简单的属性页代码：

```C#
```

## 1

```C#
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System;
using System.Runtime.InteropServices;
using System.Diagnostics;

[ComVisibleAttribute(true)]
public class clsPropMgr : PropertyManagerPage2Handler9
{

}
```

## 2控件对象定义

控件ID定义，这里的ID对后面事件是有用的。


```C#
   //属性管理器页面所需的控件对象Control objects required for the PropertyManager page 
PropertyManagerPage2 pm_Page;
PropertyManagerPageGroup pm_Group;
PropertyManagerPageSelectionbox pm_Selection;
PropertyManagerPageSelectionbox pm_Selection2;
PropertyManagerPageLabel pm_Label;
PropertyManagerPageCombobox pm_Combo;
PropertyManagerPageListbox pm_List;
PropertyManagerPageNumberbox pm_Number;
PropertyManagerPageOption pm_Radio;
PropertyManagerPageSlider pm_Slider;
PropertyManagerPageTab pm_Tab;
PropertyManagerPageButton pm_Button;
PropertyManagerPageBitmapButton pm_BMPButton;
PropertyManagerPageBitmapButton pm_BMPButton2;
PropertyManagerPageBitmap pm_Bitmap;
PropertyManagerPageActiveX pm_ActiveX;
    
//页面中的每个控件都需要一个惟一的IDEach control in the page needs a unique ID 
const int GroupID = 1;
const int LabelID = 2;
const int SelectionID = 3;
const int ComboID = 4;
const int ListID = 5;
const int Selection2ID = 6;
const int NumberID = 7;
const int RadioID = 8;
const int SliderID = 9;
const int TabID = 10;
const int ButtonID = 11;
const int BMPButtonID = 12;
const int BMPButtonID2 = 13;
const int BitmapID = 14;
const int ActiveXID = 15;
    
public void Show() { pm_Page.Show2(0); }
```

## 创建页面

Create the PropertyManager page

```C#

//属性管理器按钮类型
int options = 
(int)swPropertyManagerButtonTypes_e.swPropertyManager_OkayButton + (int)swPropertyManagerButtonTypes_e.swPropertyManager_CancelButton + (int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_LockedPage + (int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_PushpinButton;

pm_Page = (PropertyManagerPage2)swApp.CreatePropertyManagerPage("属性页名称", (int)options, this, ref longerrors);
```

### 菜单栏

在前面属性页pm_Page的基础上，AddTab添加菜单栏。

```C#
pm_Tab = pm_Page.AddTab(TabID, "菜单名", @"\图标路径.png", 0);
```

### 控件组

```C#
pm_Group = pm_Tab.AddGroupBox(GroupID, "名称", options);
```

### 控件

Controls 添加控件的方法如下：默认的对齐方式和选项，可按内容来，其他情况请<a href="[AddControl2 Method (IPropertyManagerPageGroup) - 2018 - SOLIDWORKS API Help](https://help.solidworks.com/2018/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IPropertyManagerPageGroup~AddControl2.html)">参考手册</a>

```C#
System.object AddControl2( 
   System.int ID, //对象ID，似乎在事件有用
   System.short ControlType,//控件类型(按钮，文本等)
   System.string Caption, //控件名称
   System.short LeftAlign, //控件对齐
   System.int Options,  //选项
   System.string Tip //提示
)
//默认对齐
int alignment = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_Indent;
//默认选项
int options = (int)swAddControlOptions_e.swControlOptions_Visible + (int)swAddControlOptions_e.swControlOptions_Enabled;
```

添加示例：

```C#
PropertyManagerPageSelectionbox pm_Selection;
const int SelectionID = 3;

pm_Selection = pm_Group.AddControl2(ID, (short)controlType, “名称”, (short)alignment, (int)options, “提示”);
```



类型：

| Member                                  | Description |
| :-------------------------------------- | :---------- |
| **swControlType_ActiveX**               | 10          |
| **swControlType_Bitmap**                | 14          |
| **swControlType_BitmapButton**          | 11          |
| **swControlType_Button**                | 3           |
| **swControlType_CheckableBitmapButton** | 12          |
| **swControlType_Checkbox**              | 2           |
| **swControlType_Combobox**              | 7           |
| **swControlType_Label**                 | 1           |
| **swControlType_Listbox**               | 6           |
| **swControlType_Numberbox**             | 8           |
| **swControlType_Option**                | 4           |
| **swControlType_Selectionbox**          | 9           |
| **swControlType_Slider**                | 13          |
| **swControlType_Textbox**               | 5           |
| **swControlType_WindowFromHandle**      | 15          |



## 事件

<a href="https://help.solidworks.com/2018/english/api/swpublishedapi/SolidWorks.Interop.swpublished~SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9_members.html">IPropertyManagerPage2Handler9 Member事件方法</a>

### 打开属性页时
```C#
int IPropertyManagerPage2Handler9.OnActiveXControlCreated(int Id, bool Status)
{
    Debug.Print("ActiveX control created");return 0;
}
```

### 确认和关闭时
```C#
void IPropertyManagerPage2Handler9.OnClose(int Reason)
{
    if (Reason == (int)swPropertyManagerPageCloseReasons_e.swPropertyManagerPageClose_Cancel)
    {//关闭属性页操作
Debug.Print("Cancel button clicked");
    }
    else if (Reason == (int)swPropertyManagerPageCloseReasons_e.swPropertyManagerPageClose_Okay)
    {//确认属性页操作
Debug.Print("OK button clicked");
    }
}
```

### 按钮按下时

OnButtonPress(int id)通过判断ID，来选择使用那个内容。这里多按钮的话，建议可以用switch判断。

```C#
 public void OnButtonPress(int id)
 {
     if (id == UserPMPage.buttonID1) { }
     else if (id == UserPMPage.buttonID2)    { }
 }
```

