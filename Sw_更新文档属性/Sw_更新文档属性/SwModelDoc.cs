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
            if (swApp != null)
            {
                return swApp;
            }
            else
            {
                Debug.Print("connect to solidworks on " + DateTime.Now);
                try
                {
                    swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                }
                catch (COMException)
                {
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
                                        MessageBox.Show(" 请先打开SOLIDWORKS程序 ", "SolidWorks", (MessageBoxButtons)MessageBoxButtons.OK, (MessageBoxIcon)MessageBoxIcon.Hand);
                                        swApp = null; return null;
                                    }
                                }
                            }
                        }
                    }
                }
                swApp.CommandInProgress = true;//告诉SW现在是用外部程序调用命令（优化性能）
                return swApp;
            }
        }
        //public static ModelDoc2 swDoc = ConnectToSolidWorks().ActiveDoc;


    }
}
