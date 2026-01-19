---
title: SOLIDWORKS-API
tag: SOLIDWORKS
toc: true
---

# SOLIDWORKS-API



## 对象拓扑图：



## 技能树





### 获得程序

（ISldWorks）

```C#
public static SldWorks swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");

如果是有多个版本的SW，可能需要区分：
(SldWorks)Marshal.GetActiveObject("SldWorks.Application.26");//2018
(SldWorks)Marshal.GetActiveObject("SldWorks.Application.27");//2019
(SldWorks)Marshal.GetActiveObject("SldWorks.Application.28");//2020……
```

### 获得文件

1、打开方式

```C#
swApp.OpenDoc(filename, (int)swDocumentTypes_e.swDocPART); //Fun（文件路径，文件类型）：
```

2、激活当前文件：

```C#
PartDoc swPart = (AssemblyDoc)swApp.ActiveDoc;//类型1：零件
AssemblyDoc swAssy = (AssemblyDoc)swApp.ActiveDoc; //类型2：装配体
DrawingDoc swDraw = (AssemblyDoc)swApp.ActiveDoc; //类型3：工程图
```

### 获得属性

（CustomPropertyManager）

```C#
//获得自定义属性对象（配置属性在双引号填“配置名称”）
CustomPropertyManager cusPropMgr = swDoc.Extension.CustomPropertyManager[""];

Get：
//GetAll获得自定义属性内容，将obj转属性名数组
cusPropMgr.GetAll2(ref vPropNamesObject, ref vPropTypes, ref vPropValues, swCustomInfoGetResult_e.swCustomInfoGetResult_NotPresent);
object[] vPropNames = (object[])vPropNamesObject; if (vPropNames == null) { return; }

Set：
增：Add(属性名,类型,属性值,添加时的设置)
cusPropMgr.Add3(PropertyName, (int)swCustomInfoType_e.swCustomInfoText, PropertyValue, (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);//输入属性内容
改：
vPropNames[i]=””
删：
cusPropMgr.Delete2((string)vPropNames[i]);
```

### 获得特征

遍历特征对象，根据特征名称找到特征对象来操作。

```C#
/// <summary>
/// 遍历特征对象
/// </summary>
public static void TraverseFeatures(bool isTopLevel)
        {
            //获得第一个特征，并赋值到当前特征curFeat
            Feature thisFeat = (Feature)swDoc.FirstFeature();
            Feature curFeat = default(Feature); curFeat = thisFeat;

            //当前特征非空就继续输出特征信息
            while ((curFeat != null))
            {
                Debug.Print(curFeat.Name);
                //特征操作
                bool isShowDimension = false; if (isShowDimension == true) ShowDimensionForFeature(curFeat);
    
                /*//子特征的遍历
                Feature subfeat = default(Feature);
                subfeat = (Feature)curFeat.GetFirstSubFeature();
                while ((subfeat != null))
                {
                    //if (isShowDimension == true) ShowDimensionForFeature(subfeat);
                    TraverseFeatures(subfeat, false);
                    Feature nextSubFeat = default(Feature);
                    nextSubFeat = (Feature)subfeat.GetNextSubFeature();
                    subfeat = nextSubFeat; nextSubFeat = null;
                }
                subfeat = null;*/
    
                //进入下一个特征
                Feature nextFeat = default(Feature);
                if (isTopLevel) { nextFeat = (Feature)curFeat.GetNextFeature(); }
                else { nextFeat = null; }
    
                curFeat = nextFeat; nextFeat = null;
            }
        }
```

### 获得零部件

dynamic IAssemblyDoc.GetComponents(bool ToplevelOnly);
通常用object[]接受，可在实际使用时转其他可用类型，常见转string[]。

举例：
```C#
public static string[] GetComps()
        {
            //打开文件&获得零部件集合
            AssemblyDoc swAssy = (AssemblyDoc)swApp.ActiveDoc; 
            string[] Comppaths = (string[])swAssy.GetComponents(true); return Comppaths;
        }
```

### 获得图纸
```C#
public static void GetSheetNames()
        {
            Sheet drwSheet = (Sheet)swDraw.GetCurrentSheet();       //获取当前工程图对象
            object[] sheetNames = (object[])swDraw.GetSheetNames(); //获取当前工程图中的所有图纸名称
            object[] views = (object[])drwSheet.GetViews();         //获取当前工程图中的所有图纸视图

            foreach (View view in views)//遍历工程图零部件,输入选择视图,输出零部件名
            {
                DrawingComponent comp = view.RootDrawingComponent; Debug.Print(comp.Name); //选择视图激活
                object[] childrencomps = (object[])comp.GetChildren();//获得子件对象
                //遍历工程图零部件
                for (int i = childrencomps.GetLowerBound(0); i <= childrencomps.GetUpperBound(0); i++)
                {
                    Debug.Print("零部件是" + ((DrawingComponent)childrencomps[i]).Name);
                }
            }
        }
```

# 常用操作

## 多选尺寸修改操作：

```

```





# 宏程序

尺寸标注



## Sw_实体转换钣金

solidwork程序在零件类型里有个“钣金”类型。这个钣金零件会有一些特殊的属性。如果是通过其他格式（如stp，x_t等）打开的模型，软件默认只会识别成一个实体特征。实体特征将不会带有钣金的属性。

- 钣金参数
- 钣金展开
- ......

使用我去找了相关程序来实现：实体特征转钣金的方法。



# 其他

## I版本与非I版本区别

SOLIDWORKS API中的方法、属性和对象（接口）有两个版本可供选择：

- 以**I**开头的版本（例如ISldWorks、IModelDoc2、IAnnotation、ISldWorks::IActiveDoc）
- 不以**I**开头的版本（例如SldWorks、ModelDoc2、Annotation、SldWorks::ActiveDoc）

这两个版本对应的是同一个对象或方法。主要区别如下：

- I版本的方法不会公开事件

下面是以*SldWorks*声明的变量的可用成员的快照。这些成员包含了事件。

- I版本的方法通常返回类型安全的接口版本，而不是对象或IDispatch。这意味着在编译时强制执行类型安全的语言（如C#）中不需要显式转换：

```cs
ISldWorks app;
...
IModelDoc2 model = app.IActiveDoc; //正确
IModelDoc2 model = app.ActiveDoc; //编译错误
IModelDoc2 model = app.ActiveDoc as IModelDoc2; //正确
```



参考：[SOLIDWORKS API方法及接口的I版本与非I版本的区别 | solidworks-GPT](https://jiaqiwang969.github.io/solidworks-GPT/zh-Hans/docs/codestack/solidworks-api/troubleshooting/macros/windows-api-functions-incorrect-use/i-api-versions/)
