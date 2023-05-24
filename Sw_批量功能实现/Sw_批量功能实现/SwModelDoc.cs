using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace Sw_toolkit
{
    class SwModelDoc
    {
        //封拆箱控制swApp的读写性
        public static ISldWorks swApp { get; private set; }
        //获取当前运行的程序,利用try-catch检查程序对象和SW版本。只装一个版本的可以直接GetActiveObject("SldWorks.Application")        
        public static ISldWorks ConnectToSolidWorks()
        {
            string str1 = "SldWorks.Application";
            string str2 = "SldWorks.Application";
            if (swApp != null) { return swApp; }
            else
            {
                for (int i = 20; i < 100; i++)
                {
                    str2 = str1 + "." + i; Console.WriteLine(str2);
                    try { swApp = (SldWorks)Marshal.GetActiveObject(str2); return swApp; }
                    catch (COMException) { swApp = null; }
                }
                MessageBox.Show(" 请先打开SOLIDWORKS程序 ");
                swApp.CommandInProgress = true;//告诉SW现在是用外部程序调用命令（优化性能）
                return swApp;
            }
        }
        public static ModelDoc2 swDoc = ConnectToSolidWorks().ActiveDoc;

        //打开文件
        public static ModelDoc2 OpenDoc(string str)
        {
            string postfix = str.Substring(str.LastIndexOf('.'));

            if (postfix == ".sldprt" || postfix == ".SLDPRT")
            {
                swDoc = SwModelDoc.swApp.OpenDoc(str, 1);
            }
            if (postfix == ".sldasm" || postfix == ".SLDASM")
            {
                swDoc = SwModelDoc.swApp.OpenDoc(str, 2);
            }
            if (postfix == ".slddrw" || postfix == ".SLDDRW")
            {
                swDoc = SwModelDoc.swApp.OpenDoc(str, 3);
            }
            return swDoc;
        }

        //零件配置选项设置
        public static void Macro1()
        {

            ISldWorks swApp = SwModelDoc.ConnectToSolidWorks(); if (swApp == null) { return; }
            //  if (textBox1.Text == "请选择文件" || textBox1.Text == "" || textBox2.Text == "请选择文件" || textBox2.Text == "") { MessageBox.Show(@"请选择文件"); return; }

            ModelDoc2 swDoc = ((ModelDoc2)(swApp.ActiveDoc));
            ModelView myModelView = null;
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.FrameState = ((int)(swWindowState_e.swWindowMaximized));
            swDoc.Extension.SelectByID2("默认", "CONFIGURATIONS", 0, 0, 0, false, 0, null, 0);
            swDoc.EditConfiguration3("默认", "默认", " ", " ", 9181);

            IConfiguration config = swDoc.GetConfigurationByName("默认");

            config.ChildComponentDisplayInBOM = 1;
        }

        //
        public static void Mainxx()
        {
            //参考SOLIDWORKS API Help中的IExportToDWG2 Method (IPartDoc) 和 Export Part to DWG Example (C#)
            //首先在钣金中绘制三维草图，长边代表导出dxf图的X方向，短边代表dxf图的Y方向，用于限定钣金拉丝方向(默认是X方向)，防止排版错误

            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc; //获取当前已打开的零件
            if (swModel != null)
            {
                PartDoc swPart = (PartDoc)swModel;
                string swModelName = swModel.GetPathName();
                string swDxfName = swModelName.Substring(0, swModelName.Length - 6) + "dxf";
                double[] dataAlignment = new double[12];
                dataAlignment[0] = 0.0;
                dataAlignment[1] = 0.0;
                dataAlignment[2] = 0.0;
                dataAlignment[3] = 1.0;
                dataAlignment[4] = 0.0;
                dataAlignment[5] = 0.0;
                dataAlignment[6] = 0.0;
                dataAlignment[7] = 1.0;
                dataAlignment[8] = 0.0;
                dataAlignment[9] = 0.0;
                dataAlignment[10] = 0.0;
                dataAlignment[11] = 1.0;
                //Array[0], Array[1], Array[2] - XYZ coordinates of new origin
                //Array[3], Array[4], Array[5] - coordinates of new x direction vector
                //Array[6], Array[7], Array[8] - coordinates of new y direction vector
                //判断XYAXIS，长边作为X轴，短的作为Y轴，用于限定拉丝方向
                bool status = swModel.Extension.SelectByID2("XYAXIS", "SKETCH", 0, 0, 0, false, 0, null, 0);
                if (status)
                {
                    SelectionMgr swSelectionMgr = swModel.SelectionManager;
                    Feature swFeature = swSelectionMgr.GetSelectedObject6(1, -1);
                    Sketch swSketch = swFeature.GetSpecificFeature2();
                    var swSketchPoints = swSketch.GetSketchPoints2();//获取草图中的所有点
                                                                     //用这三个点抓取直线，并判断长度，长边作为X轴，画3D草图的时候一次性画出两条线，不能分两次画出，否则会判断错误
                    SketchPoint p0 = swSketchPoints[0];//最先画的点
                    SketchPoint p1 = swSketchPoints[1];//作为坐标原点
                    SketchPoint p2 = swSketchPoints[2];//最后画的点
                    dataAlignment[0] = p1.X * 1000;
                    dataAlignment[1] = p1.Y * 1000;
                    dataAlignment[2] = p1.X * 1000;
                    double l1 = Math.Sqrt(Math.Pow(p0.X - p1.X, 2) + Math.Pow(p0.Y - p1.Y, 2) + Math.Pow(p0.Z - p1.Z, 2));
                    double l2 = Math.Sqrt(Math.Pow(p2.X - p1.X, 2) + Math.Pow(p2.Y - p1.Y, 2) + Math.Pow(p2.Z - p1.Z, 2));
                    if (l1 > l2)
                    {
                        dataAlignment[3] = p0.X * 1000 - p1.X * 1000;
                        dataAlignment[4] = p0.Y * 1000 - p1.Y * 1000;
                        dataAlignment[5] = p0.Z * 1000 - p1.Z * 1000;
                        dataAlignment[6] = p2.X * 1000 - p1.X * 1000;
                        dataAlignment[7] = p2.Y * 1000 - p1.Y * 1000;
                        dataAlignment[8] = p2.Z * 1000 - p1.Z * 1000;
                    }
                    else
                    {
                        dataAlignment[3] = p2.X * 1000 - p1.X * 1000;
                        dataAlignment[4] = p2.Y * 1000 - p1.Y * 1000;
                        dataAlignment[5] = p2.Z * 1000 - p1.Z * 1000;
                        dataAlignment[6] = p0.X * 1000 - p1.X * 1000;
                        dataAlignment[7] = p0.Y * 1000 - p1.Y * 1000;
                        dataAlignment[8] = p0.Z * 1000 - p1.Z * 1000;
                    }
                }
                object varAlignment = dataAlignment;

                //Export sheet metal to a single drawing file将钣金零件导出单个dxf文件
                //include flat-pattern geometry，倒数第二位数字1代表钣金展开，options = 1;
                swPart.ExportToDWG2(swDxfName, swModelName, (int)swExportToDWG_e.swExportToDWG_ExportSheetMetal, true, varAlignment, false, false, 4095, null);
            }

        }

        public static void ExportToDWG2(ModelDoc2 swDoc)
        {
            //获取当前已打开的零件
            PartDoc swPart = (PartDoc)swDoc;
            //获取保存路径
            string swModelName = swDoc.GetPathName();
            string swDxfName = swModelName.Substring(0, swModelName.Length - 6) + "dxf";
            //double[] dataAlignment = new double[12]; object varAlignment = dataAlignment;

            #region
            /*
            FilePath
            导出的 DXF/DWG 文件的路径和文件名
            ModelName
            活动部件文档的路径和文件名
            Action
            swExportToDWG_e中定义的导出操作
            ExportToSingleFile
            如果为 True，则另存为一个文件;假以另存为多个文件
            Alignment 
            进程内非托管C++：指向包含与输出对齐相关的信息的 12 个双精度值数组的指针（请参阅备注) 
            VBA、VB.NET、C# 和 C++/CLI：不支持                
            IsXDirFlipped
            如果为 true，则翻转 x 方向; 否则为假
             IsYDirFlipped
            如果为 true，则翻转 y 方向; 否则为假
             SheetMetalOptions
            包含钣金件导出选项的位掩码; 仅当操作 = swExportToDWG_e时有效.swExportToDWG_ExportSheetMetal（请参阅备注)
            ViewsCount
            要导出的注释视图的数量; 仅当操作 = swExportToDWG_e 时才有效.swExportToDWG_ExportAnnotationViews
             Views
            进程内非托管C++：指向要导出的注释视图名称数组的指针; 仅当操作 = swExportToDWG_e 时才有效.swExportToDWG_ExportAnnotationViews
             VBA、VB.NET、C# 和 C++/CLI：不支持
            有关此类方法的详细信息，请参阅进程内方法。
            */
            #endregion
            swPart.ExportToDWG2(swDxfName, swModelName, (int)swExportToDWG_e.swExportToDWG_ExportSheetMetal, true, null, false, false, 2149, null);


        }

        
    }
}
