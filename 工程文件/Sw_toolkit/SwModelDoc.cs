using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms;
using MessageBox = System.Windows.MessageBox;

namespace Sw_toolkit
{
    class SwModelDoc
    {
        public static ISldWorks swApp { get; private set; }
        /// <summary>
        /// //获取当前运行的程序,利用try-catch检查程序对象和SW版本。只装一个版本的可以直接GetActiveObject("SldWorks.Application")
        /// </summary>
        /// <returns></returns>
        public static ISldWorks ConnectToSolidWorks()
        {
            if (swApp != null)
            {
                return swApp;
            }
            else
            {
                Debug.Print("connect to solidworks on " + DateTime.Now);
                try
                {
                    swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application.26");//SW2018
                }
                catch (COMException)
                {
                    try
                    {
                        swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application.27");//SW2019
                    }
                    catch (COMException)
                    {
                        try
                        {
                            swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application.28");//SW2020
                        }
                        catch (COMException)
                        {
                            try
                            {
                                swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application.29");//SW2021
                            }
                            catch (COMException)
                            {
                                try
                                {
                                    swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application.30");//SW2022
                                }
                                catch (COMException)
                                {
                                    MessageBox.Show(" 请先打开SOLIDWORKS程序 ", "SolidWorks", (MessageBoxButton)MessageBoxButtons.OK, (MessageBoxImage)MessageBoxIcon.Hand);
                                    swApp = null;
                                }
                            }
                        }
                    }
                }
                swApp.CommandInProgress = true;//告诉SW现在是用外部程序调用命令（优化性能）
                return swApp;
            }
        }
        public static ModelDoc2 swModel = (ModelDoc2)ConnectToSolidWorks().ActiveDoc;

        /// <summary>
        /// 打开文件(包括零件、装配体、工程图)
        /// </summary>
        /// <param name="filename"></param>
        public static void Open(string filename)
        {
            try { swApp.OpenDoc(filename, (int)swDocumentTypes_e.swDocPART); }
            catch (Exception)
            {
                try
                { swApp.OpenDoc(filename, (int)swDocumentTypes_e.swDocASSEMBLY); }
                catch (Exception)
                { swApp.OpenDoc(filename, (int)swDocumentTypes_e.swDocDRAWING); }
                return;
            }


        }

        /// <summary>
        /// 删除全部自定义属性 
        /// </summary>
        /// <param name="filename"></param>
        public static void DelProperties(string filename)
        {
            object vPropNamesObject = null;
            object vPropTypes = null;
            object vPropValues = null;
            swModel = (ModelDoc2)swApp.ActiveDoc;
            //获得自定义属性对象
            CustomPropertyManager cusPropMgr = swModel.Extension.CustomPropertyManager[""];
            //GetAll获得自定义属性内容
            cusPropMgr.GetAll2(ref vPropNamesObject, ref vPropTypes, ref vPropValues, swCustomInfoGetResult_e.swCustomInfoGetResult_NotPresent);
            object[] vPropNames = (object[])vPropNamesObject;
            // 对于每个自定义属性，删除操作
            if (vPropNames == null) { return; }
            for (int i = 0; i < vPropNames.Length; i++)
            {
                string PropName = (string)vPropNames[i];
                cusPropMgr.Delete2(PropName);
                Debug.WriteLine(PropName);
            }
            swModel.Save(); swModel.Close();
        }

        /// <summary>
        /// 删除指定自定义属性 
        /// </summary>
        /// <param name="PropNames"></param>
        public static void DelTheProperties(string PropNames)
        {
            string[] vPropNames = PropNames.Split('，');
            CustomPropertyManager cusPropMgr = swModel.Extension.CustomPropertyManager[""];
            for (int i = 0; i <= vPropNames.Length; i++) { cusPropMgr.Delete2(vPropNames[i]); }
            swModel.Save(); swModel.Close();
        }

        public static void GetSelecte(string PropNames)//获得所选择的对象
        {
            swModel.Extension.SelectByID2("手推车-53@工程图视图2/支撑柱-1@手推车", "COMPONENT", 0, 0, 0, false, 0, null, 0);
            SelectionMgr swSelMgr = (SelectionMgr)swModel.SelectionManager;
            object swDrawComp = swSelMgr.GetSelectedObjectsComponent4(1, 0);

        }

        public static void ing()//测试
        {

               SelectionMgr swSelectionMgr;

            bool selectionRet = swModel.Extension.SelectByID2("blade shaft-1@98food processor/drive shaft pin-1@blade shaft", "COMPONENT", 0, 0, 0, false, 0, null, 0);
            swSelectionMgr = (SelectionMgr)swModel.SelectionManager;
            if (selectionRet)
            {
                Component2 selectedComponent;
                selectedComponent = (Component2)swSelectionMgr.GetSelectedObjectsComponent2(1);
                //selectedComponent.SetSuppression(1);
                ModelDoc2 componentModel;
                componentModel = (ModelDoc2)selectedComponent.GetModelDoc();
                Debug.Print(componentModel.GetTitle());
                string a = selectedComponent.GetPathName();



            }


        }

    }

    class SwModelDoc_Check
    {
        public static SldWorks swApp = (SldWorks)SwModelDoc.ConnectToSolidWorks();
        public static void Checkopen1()         // 启动程序前，检查是否打开零件:
        {
            if (swApp != null)
            {
                ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
                if (swModel == null) MessageBox.Show("请打开模型文件");
            }
            else MessageBox.Show("请打开软件");
        }
        public static void RevisionNumber()     //显示版本号
        {
            swApp.CommandInProgress = true;//告诉SW现在是用外部程序调用命令（优化性能）
            string msg = "SOLIDWORKS 的版本是 " + swApp.RevisionNumber();
            MessageBox.Show(msg, "版本信息", MessageBoxButton.OK);
        }
    }
}
