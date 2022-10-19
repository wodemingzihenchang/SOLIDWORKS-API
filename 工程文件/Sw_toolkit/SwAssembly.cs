using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Diagnostics;
using System.Windows;

namespace Sw_toolkit
{
    class SwAssembly
    {
        public static SldWorks swApp = (SldWorks)SwModelDoc.ConnectToSolidWorks();//获取当前运行的程序 
        public static ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;


        /// <summary>
        /// 遍历装配体组件对象数组
        /// </summary>
        /// <param name="swAssyname"></param>
        /// <returns>零部件数组Comppaths[i]</returns>
        public static string[] GetComps()
        {
            AssemblyDoc swAssy = (AssemblyDoc)swApp.ActiveDoc; //打开文件
            object[] Comps = (object[])swAssy.GetComponents(true);         //获得零部件集合
            string[] Comppaths = new string[Comps.Length];

            swApp.CommandInProgress = true;                                 //加速遍历
            for (int i = 0; i < Comps.Length; i++)                          //循环处理每个零部件
            {
                Component Comp = (Component)Comps[i];
                Debug.Print("零部件名字：" + Comp.Name);
                Comppaths[i] = Comp.GetPathName();                          //填入数组
            }
            swApp.CommandInProgress = false;
            return Comppaths;
        }
        public static string[] GetComps(string swAssyname)
        {
            AssemblyDoc swAssy = (AssemblyDoc)swApp.OpenDoc(swAssyname, 2); //打开文件
            object[] Comps = (object[])swAssy.GetComponents(false);         //获得零部件集合
            string[] Comppaths = new string[Comps.Length];

            swApp.CommandInProgress = true;                                 //加速遍历
            for (int i = 0; i < Comps.Length; i++)                          //循环处理每个零部件
            {
                Component Comp = (Component)Comps[i];
                Debug.Print("零部件名字：" + Comp.Name);
                Comppaths[i] = Comp.GetPathName();                          //填入数组
            }
            swApp.CommandInProgress = false;
            return Comppaths;
        }

        /// <summary>
        /// 设置材料明细表,输入表名，配置
        /// </summary>
        /// <param name="bomname"></param>
        /// <param name="config"></param>
        public static void SetBOM(string bomname, string config)//
        {
            AssemblyDoc swAssy = (AssemblyDoc)swApp.ActiveDoc;
            Feature f = (Feature)swAssy.FeatureByName(bomname);
            BomFeature bom = (BomFeature)f.GetSpecificFeature();//获取BOM表,吧特征f转成特殊特征类（bom是属于API的特殊特征类）

            bom.Configuration = config;                         //设置BOM表
            string a = bom.Configuration;
            Debug.Print(f.Name + "的配置是" + a);               //测试
        }

        /// <summary>
        /// 装配体零件改名
        /// </summary>
        /// <param name="oldname"></param>
        /// <param name="newname"></param>
        public static void Rename(string oldname, string newname)
        {
            //装配体选择子装配体零件要特定名字，搞几个名字变量存储信息
            string Name = null, Name0 = null, Name1 = null;

            //获得顶层组件——遍历组件
            AssemblyDoc swAssy = (AssemblyDoc)swApp.ActiveDoc;
            object[] Comps = (object[])swAssy.GetComponents(false);
            for (int i = 0; i < Comps.Length; i++)
            {
                Name = null; Name0 = null; Name1 = null;
                Component2 component2 = (Component2)Comps[i];

                //处理路径名字格式，有点类似构造函数？？？
                string[] Names = component2.Name.Split('/');
                for (int ii = 0; ii < Names.Length; ii++)
                {
                    Name0 = Names[0] + "@1/";
                    if (ii >= 1) { Name1 += Names[ii] + "@" + Names[ii - 1].Substring(0, Names[ii - 1].LastIndexOf('-')) + "/"; }
                }
                Name = Name0 + Name1; Debug.WriteLine(Name);

                //获得新名字
                string NewName = null;
                NewName = component2.Name.Substring(0, component2.Name.LastIndexOf('-')).Replace(oldname, newname);
                if (NewName.LastIndexOf('/') != -1) { Debug.WriteLine(NewName.LastIndexOf('/')); NewName = NewName.Substring(NewName.LastIndexOf('/') + 1, NewName.Length - 1 - NewName.LastIndexOf('/')); }

                //选择重命名,
                ModelDocExtension swModelDocExt = swModel.Extension;
                swModelDocExt.SelectByID2(Name, "COMPONENT", 0, 0, 0, false, 0, null, 0);
                swModelDocExt.RenameDocument(NewName); swModel.ClearSelection2(true);
            }
        }


        //找出不在同文件夹位置的工程图
        public static void FindDrawing()
        {
            //获得零部件集合
            AssemblyDoc swAssy = (AssemblyDoc)swApp.ActiveDoc;
            object[] Comps = (object[])swAssy.GetComponents(false);

            for (int i = 0; i < Comps.Length; i++)
            {
                Component2 Comp = (Component2)Comps[i];


            }


        }
        //打开工程图
        public static void OpenDrawings(string[] vPaths)
        {
            if (vPaths != null)
            {
                for (int i = 0; i < vPaths.Length; i++)
                {
                    string drwFilePath;
                    drwFilePath = vPaths[i];

                    DocumentSpecification swDocSpec = (DocumentSpecification)swApp.GetOpenDocSpec(drwFilePath);
                    ModelDoc2 swDraw = swApp.OpenDoc7(swDocSpec);
                }
            }
        }

    }
}





/*//找关联的工程图
public static void FindAssociatedDrawings(rootDir As String, filePath As String)
{
    string paths();
    object fso = CreateObject("Scripting.FileSystemObject");

    object folder = fso.GetFolder(rootDir);


    CollectDrawingFilesFromFolder(folder, filePath, paths);


If(Not paths) <> -1 Then
FindAssociatedDrawings = paths
Else
Err.Raise vbError, "", "Failed to find the associated drawings for " & filePath
End If

End Function

Sub CollectDrawingFilesFromFolder(folder As Object, targetFilePath As String, ByRef paths() As String)


For Each file In folder.files

Dim fileExt As String
fileExt = Right(file.path, Len(file.path) - InStrRev(file.path, "."))


If LCase(fileExt) = LCase("slddrw") Then

    If IsReferencingDrawing(file.path, targetFilePath) Then
        If(Not paths) = -1 Then
           ReDim paths(0)
        Else
            ReDim Preserve paths(UBound(paths) +1)
        End If
        paths(UBound(paths)) = file.path
    End If
End If
Next


Dim subFolder As Object
For Each subFolder In folder.SubFolders
CollectDrawingFilesFromFolder subFolder, targetFilePath, paths
Next
}
*/


/*获得装配体零件名+替换同名的零件
public static void TraverseComponentq(Component2 swComp, long nLevel)
{
    //激活文件
    ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
    Configuration swConf = (Configuration)swModel.GetActiveConfiguration();
    Component2 swRootComp = (Component2)swConf.GetRootComponent();

    Component2 swChildComp;
    Component2 swtemp;
    string sPadStr = " ";

    for (int i = 0; i <= nLevel - 1; i++)
    {
        sPadStr = sPadStr + " ";
    }

    object[] vChildComp = (object[])swComp.GetChildren();
    for (int i = 0; i < vChildComp.Length; i++)
    {
        swChildComp = (Component2)vChildComp[i];
        Debug.Print(swChildComp.Name2);
        for (int j = 0; j < vChildComp.Length; j++)
        {
            swtemp = (Component2)vChildComp[j];
            Debug.Print(swtemp.Name2);
            if (File.ChecksameName(swChildComp.Name2) + "-1" == swtemp.Name2 && File.ChecksameName2(swChildComp.Name2) != -1)
            {
                swModel = (ModelDoc2)swApp.ActiveDoc;
                ModelDocExtension swModelDocExt = (ModelDocExtension)swModel.Extension;
                SelectionMgr selectionMgr = (SelectionMgr)swModel.SelectionManager;
                AssemblyDoc assemblyDoc = (AssemblyDoc)swModel;
                //选择零部件
                swModelDocExt.SelectByID2(swChildComp.Name2 + "@" + Path.GetFileNameWithoutExtension(swModel.GetPathName()), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                //替代零部件
                assemblyDoc.ReplaceComponents(swtemp.GetPathName(), "", false, true);
            }
        }
    }
}*/





