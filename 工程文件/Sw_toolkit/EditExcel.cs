using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using MessageBox = System.Windows.Forms.MessageBox;

namespace Sw_toolkit
{
    class EditExcel
    {
        // 各种对象的定义
        public static Excel.Application excel = null;   //程序对象
        public static Excel.Workbooks books = null;     //工作簿组
        public static Excel.Workbook book = null;       //工作簿
        public static Excel.Sheets sheets = null;       //工作表组
        public static Excel.Worksheet sheet = null;     //工作表
        public static Excel.Range cells = null;         //单元网格组
        public static Excel.Range range = null;         //范围
        public static int loading = 0;                  //进度条


        public static void GetPropertiesToExcel()//获取自定义属性到Excel
        {
            SldWorks swApp = (SldWorks)SwModelDoc.ConnectToSolidWorks();//获取当前运行的程序
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            object vPropNamesObject = null;
            object vPropTypes = null;
            object vPropValues = null;
            //获得自定义属性对象
            CustomPropertyManager cusPropMgr = swModel.Extension.CustomPropertyManager[""];
            //获得所有配置名称,数量
            string[] ConfigNames = (string[])swModel.GetConfigurationNames();
            Debug.Print("此配置的自定义属性数量:" + cusPropMgr.Count);
            //获取自定义属性内容，ref用以返回属性名，类型，值
            cusPropMgr.GetAll2(ref vPropNamesObject, ref vPropTypes, ref vPropValues, swCustomInfoGetResult_e.swCustomInfoGetResult_NotPresent);
            object[] vPropNames = (object[])vPropNamesObject;
            string[] propValues = (string[])vPropValues;

            NewExcel();//新建空白Excel表，Excel对象初始化，属性值填入表格
            excel = new Excel.Application();
            books = excel.Workbooks;
            try
            {
                book = books.Open(@"C:\temp\属性.xlsx");    // 打开文件                         
                sheets = book.Worksheets;   // 选中文件里的所有表                      
                sheet = sheets[1];          // 选中第一个表                        
                cells = sheet.Cells;        // 选中第一个表的所有单元格
                cells[1, 1] = "名称";
                cells[1, 2] = "路径";
                for (int i = 3; i < vPropNames.Length + 3; i++)
                {
                    cells[1, i] = vPropNames[i - 3];
                }
                cells[2, 1] = Path.GetFileName(swModel.GetPathName());
                cells[2, 2] = swModel.GetPathName();
                for (int i = 3; i < propValues.Length + 3; i++)
                {
                    cells[2, i] = propValues[i - 3];
                }
                book.Close(true);           // 保存文件的修改并关闭
                excel.Quit();               // 关闭Excel
            }
            finally
            {
                Marshal.FinalReleaseComObject(excel);
                GC.Collect();
            }
            MessageBox.Show("完成");
        }
        public static void GetPropnameToExcel(string filenames, int Doctype, int times)//获取自定义属性名到Excel
        {
            SldWorks swApp = (SldWorks)SwModelDoc.ConnectToSolidWorks();//获取当前运行的程序
            object vPropNamesObject = null;
            object vPropTypes = null;
            object vPropValues = null;
            //获得自定义属性对象
            ModelDoc2 swModel = (ModelDoc2)swApp.OpenDoc(filenames, Doctype);
            CustomPropertyManager cusPropMgr = swModel.Extension.CustomPropertyManager[""];
            //获取自定义属性内容，ref用以返回属性名，类型，值
            cusPropMgr.GetAll2(ref vPropNamesObject, ref vPropTypes, ref vPropValues, swCustomInfoGetResult_e.swCustomInfoGetResult_NotPresent);
            object[] vPropNames = (object[])vPropNamesObject;
            string[] propValues = (string[])vPropValues;
            //写入表头
            cells[1, 1] = "名称";
            cells[1, 2] = "路径";
            if (vPropNames == null) { return; }
            for (int ii = 3; ii < vPropNames.Length + 3; ii++)
            {
                cells[times, ii] = vPropNames[ii - 3];
            }
            swApp.CloseDoc(filenames);
        }
        public static void GetPropvalueToExcel(string filenames, int Doctype, int times)//获取自定义属性值到Excel
        {
            SldWorks swApp = (SldWorks)SwModelDoc.ConnectToSolidWorks();//获取当前运行的程序
            object vPropNamesObject = null;
            object vPropTypes = null;
            object vPropValues = null;
            //获得自定义属性对象
            ModelDoc2 swModel = (ModelDoc2)swApp.OpenDoc(filenames, Doctype);
            CustomPropertyManager cusPropMgr = swModel.Extension.CustomPropertyManager[""];
            //获取自定义属性内容，ref用以返回属性名，类型，值
            cusPropMgr.GetAll2(ref vPropNamesObject, ref vPropTypes, ref vPropValues, swCustomInfoGetResult_e.swCustomInfoGetResult_NotPresent);
            object[] vPropNames = (object[])vPropNamesObject;
            string[] propValues = (string[])vPropValues;

            cells[times, 1] = Path.GetFileName(swModel.GetPathName());
            cells[times, 2] = swModel.GetPathName();
            //写入内容
            if (vPropNames == null) { return; }
            for (int ii = 3; ii < propValues.Length + 3; ii++)
            {
                cells[times, ii] = propValues[ii - 3];
            }
            //关闭文件
            swApp.CloseDoc(filenames);
        }


        public static void SetPropertiesToExcel()//获取自定义属性到Excel
        {
            SldWorks swApp = (SldWorks)SwModelDoc.ConnectToSolidWorks();//获取当前运行的程序
            excel = new Excel.Application();                            //Excel对象初始化，属性值填入表格
            try
            {
                book = excel.Workbooks.Open(@"C:\temp\属性.xlsx");      // 打开文件
                Worksheet wsh = book.Sheets[1];                         // 获得工作表1
                int i = 2;
                object filename = wsh.Cells[i, 2].Value;                // 获得文件路径对象
                object propname = wsh.Cells[1, 3].Value;                // 获得属性内容对象
                int row = 1; object rowobj = wsh.Cells[i, 2].Value;

                for (row = 2; rowobj != null; row++) { rowobj = wsh.Cells[row + 1, 2].Value; }
                Form1 Form1 = new Form1(); Form1.loading_Initialize(row - 2); Form1.Hide(); Form1.Show();//初始化进度条

                for (i = 2; filename != null; i++)
                {
                    ModelDoc2 swModel = (ModelDoc2)swApp.OpenDoc(filename.ToString(), 1);
                    CustomPropertyManager cusPropMgr = swModel.Extension.CustomPropertyManager[""];//获得自定义属性对象
                    for (int ii = 3; propname != null; ii++)
                    {
                        string PropertyName = wsh.Cells[1, ii].Value.ToString();    // 获得属性名
                        string PropertyValue = "";
                        if (wsh.Cells[i, ii].Value != null)
                        {
                            PropertyValue = wsh.Cells[i, ii].Value.ToString();     // 获得属性值  
                        }
                        cusPropMgr.Add3(PropertyName, (int)swCustomInfoType_e.swCustomInfoText, PropertyValue, (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);//输入属性内容
                        Debug.WriteLine(PropertyName + "：" + PropertyValue);
                        propname = wsh.Cells[1, ii + 1].Value;
                    }
                    swModel.Save();
                    swApp.CloseDoc(filename.ToString());
                    filename = wsh.Cells[i + 1, 2].Value;
                    propname = wsh.Cells[1, 3].Value;
                    Form1.loading();
                }
                book.Close(true);                           // 保存文件的修改并关闭
                excel.Quit();                               // 关闭Excel
            }
            finally
            {
                //Release();
                Marshal.FinalReleaseComObject(excel);
                GC.Collect();
            }
            MessageBox.Show("完成");
            //}            else MessageBox.Show("请关闭文件，使SW处于开始界面");
        }

        public static void NewExcel()   //新建空白Excel表
        {
            //需要生成的excel完整的路径以及名称
            string saveExcelPath = @"C:\temp\属性.xlsx";
            //新建一个ExcelPackage
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage excelPkg = new ExcelPackage();
            //添加一个sheet
            ExcelWorksheet excelWorksheet = excelPkg.Workbook.Worksheets.Add("Sheet1");
            //保存这个excel
            excelPkg.SaveAs(new FileInfo(saveExcelPath));
            //初始化
            excel = new Excel.Application();
            books = excel.Workbooks;
            book = books.Open(@"C:\temp\属性.xlsx");    // 打开文件                         
            sheets = book.Worksheets;   // 选中文件里的所有表                      
            sheet = sheets[1];          // 选中第一个表                        
            cells = sheet.Cells;        // 选中第一个表的所有单元格      
        }

        public static void Write()      //写入Excel表
        {
            //Excel对象初始化
            excel = new Excel.Application();
            books = excel.Workbooks;
            string file = @"C:\Users\DELL\Desktop\新建 Microsoft Excel 工作表.xlsx";
            try
            {
                book = books.Open(file);    // 打开文件                         
                sheets = book.Worksheets;   // 选中文件里的所有表                      
                sheet = sheets[1];          // 选中第一个表                        
                cells = sheet.Cells;        // 选中第一个表的所有单元格
                range = cells[2, 1];        // 选中单元格[2,1]   //range = sheet.Cells[2, 1]这种写法会产生sheet.Cells和sheet.Cells[2, 1]这两个Excel.Range对象。其中一个对象无法得到释放，会导致Excel进程残留问题
                range.Value = "2,1";        // 将单元格[2,1]的文字设置成test
                book.Close(true);           // 保存文件的修改并关闭
                excel.Quit();               // 关闭Excel
            }
            finally
            {
                Release();
            }
        }

        public static void Release()    //释放之前定义的对象，每循环一次就要释放一次，不然会导致Excel进程残留问题
        {
            book.Close(true);           // 保存文件的修改并关闭
            excel.Quit();               // 关闭Excel
            Marshal.FinalReleaseComObject(excel);//GC.Collect();
            /*
            Marshal.FinalReleaseComObject(range);
            Marshal.FinalReleaseComObject(cells);
            Marshal.FinalReleaseComObject(sheet);
            Marshal.FinalReleaseComObject(sheets);
            Marshal.FinalReleaseComObject(book);
            Marshal.FinalReleaseComObject(books);
            Marshal.FinalReleaseComObject(excel);
            */
        }
        public static void Excelexample()//
        {
            //应用程序
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            //工作簿
            Workbook wbk = app.Workbooks.Open("");
            //工作表
            Worksheet wsh = wbk.Sheets["All"];
            //读取
            string str = wsh.Cells[1, 1].Value.ToString();
            //写入，索引以1开始
            wsh.Cells[2, 1] = "str";
            //保存
            wbk.Save();
            //退出
            app.Quit();
            //释放
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
        }

    }
}

