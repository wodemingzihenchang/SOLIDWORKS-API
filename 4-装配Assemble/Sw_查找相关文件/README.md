---
title: 查找相关文件
tag: SOLIDWORKS
toc: true
---

# 查找相关文件

查找相关文件(参考引用)

## 简介

<img src="README\功能简介1.png">

<img src="README\功能简介2.png">

在Solidworks2020的API中，我们可以通过文档扩展对象的方法获得文档的参考引用文件ModelDocExtension::GetDependencies。

```
ModelDocExtension.GetDependencies(Traverseflag, Searchflag, AddReadOnlyInfo, ListBrokenRefs, AppendImportedPaths)
```



在一些低版本的Solidworks中，可以使用文档对象的相应方法获得参考文件ModelDoc2::GetDependencies2。


```
ModelDoc2.GetDependencies2(Traverseflag, Searchflag, AddReadOnlyInfo)
```

## 代码实例：



```csharp
public static void GetDocReference(ModelDoc2 AssemDoc)
{
      //旧方法
      object[] ObjFiles1 = AssemDoc.GetDependencies2(true,false,true);
      StringBuilder Sb = new StringBuilder("ModelDoc::GetDependencies2方法:\r\n");
      foreach (object of in ObjFiles1)
      {
           Sb.Append(of.ToString().Trim() + "\r\n");
      }
      System.Windows.MessageBox.Show(Sb.ToString().Trim());
      
      //新方法
      ModelDocExtension AssemDocEx = AssemDoc.Extension;
      object[] ObjFiles2 = AssemDocEx.GetDependencies(true, false, true, true, true);
      Sb = new StringBuilder("ModelDocExtension::GetDependencies方法:\r\n");
      foreach (object of in ObjFiles2)
      {
           Sb.Append(of.ToString().Trim() + "\r\n");
      }
      System.Windows.MessageBox.Show(Sb.ToString().Trim());
}
```

## 实例效果：

因为我上面代码做了字符串拆分，所以数组的内容可能有所不同。

ModelDoc2::GetDependencies2旧方法会输出：[0]是名称，[1]路径，[]往复循环

ModelDocExtension::GetDependencies新方法会输出：[0]是名称，[1]路径，[2]是否读写，[]往复循环（如下图：）

<img src="README\实例解读.png">
