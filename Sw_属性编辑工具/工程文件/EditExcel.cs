using Microsoft.Office.Interop.Excel;
using Excel=Microsoft.Office.Interop.Excel;
using SolidWorks.Interop.swdocumentmgr;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Sw_属性编辑工具
{
    class EditExcel
    {
        #region /*————  各种对象的定义  ————*/
        public static Excel.Application excel = null;   //程序对象
        public static Workbooks books = null;     //工作簿组
        public static Workbook book = null;       //工作簿
        public static Sheets sheets = null;       //工作表组
        public static Worksheet sheet = null;     //工作表
        public static Range cells = null;         //单元网格组
        public static Range range = null;         //范围


        //需要生成的excel完整的路径以及名称
        public string  filePath = System.Windows.Forms.Application.StartupPath + "Sw_属性编辑工具.xlsx";
        #endregion

        /*————  属性操作  ————*/
        /// <summary>
        /// 输出属性-SwDocumentManager
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="Doctype">1</param>
        public static void GetPropnameToExcel_SWDM(string[] filepath)
        {
            //创建一个Excel表
            EditExcel.NewExcel(); int Row = 2;

            //遍历文件
            foreach (string filename in filepath)
            {
                //获得文件属性
                SwDMDocument20 swDoc = (SwDMDocument20)SwDocumentManager.SwDMDocument(filename);
                string[] vPropNames = (string[])swDoc.GetCustomPropertyNames();
                //写入表头
                cells[1, 1] = "名称";
                cells[1, 2] = "路径";
                //写入属性
                cells[Row, 1] = Path.GetFileName(swDoc.FullName);
                cells[Row, 2] = swDoc.FullName;
                if (vPropNames != null)
                {
                    for (int Colum = 3; Colum < vPropNames.Length + 3; Colum++)
                    {
                        string propValues = swDoc.GetCustomProperty(vPropNames[Colum - 3], out SwDmCustomInfoType nPropType);
                        cells[1, Colum] = vPropNames[Colum - 3];
                        cells[Row, Colum] = propValues;
                    }
                }
                //更新进度条
                swDoc.Save(); swDoc.CloseDoc();
                Row++;
            }
            MessageBox.Show(@"完成，Excel属性表保存在C:\WINDOWS\Temp\属性.xlsx");
            EditExcel.Release(excel); GC.Collect();
            System.Diagnostics.Process.Start(@"C:\WINDOWS\Temp\属性.xlsx");
        }
        /// <summary>
        /// 添加属性-SwDocumentManager
        /// </summary>
        public static void SetPropertiesToExcel_SWDM()
        {
            //读取Excel-Sheet1
            excel = new Excel.Application();
            try
            {
                book = excel.Workbooks.Open(@"C:\WINDOWS\Temp\属性.xlsx");
                Worksheet wsh = book.Sheets[1];
                //文件路径初始化
                SwDMDocument20 swDoc = null;
                int Row = 2; object filename = wsh.Cells[Row, 2].Value;
                int Colum = 3; object propname = wsh.Cells[1, 3].Value;
                object propvalue = null;

                //双循环批量处理行列数据
                for (Row = 2; wsh.Cells[Row, 2].Value != null; Row++)
                {
                    swDoc = (SwDMDocument20)SwDocumentManager.SwDMDocument((string)filename);
                    propname = wsh.Cells[1, 3].Value;
                    for (Colum = 3; wsh.Cells[1, Colum].Value != null; Colum++)
                    {
                        //添加属性名                    wsh.Range[1, Colum].NumberFormatLocal = "@";
                        propname = wsh.Cells[1, Colum].Value;
                        if (propname.GetType() != typeof(string)) { propname = Convert.ToString(wsh.Cells[1, Colum].Value); }

                        //添加属性值                    wsh.Range[Row, Colum].NumberFormatLocal = "@";
                        propvalue = wsh.Cells[Row, Colum].Value;
                        if (propvalue.GetType() != typeof(string)) { propvalue = Convert.ToString(wsh.Cells[Row, Colum].Value); }

                        SwDocumentManager.AddCustomProperties(swDoc, (string)propname, (string)propvalue);
                    }
                    //更新进度条
                    swDoc.Save();
                    filename = wsh.Cells[Row + 1, 2].Value;
                }
                swDoc.CloseDoc();
                // 保存文件的修改并关闭Excel
                book.Close(true); excel.Quit();
            }
            catch (COMException)
            {
                Marshal.FinalReleaseComObject(excel); GC.Collect();
                //MessageBox.Show("完成");
            }
        }


        /*———— Excel操作  ————*/
       

        //新建空白Excel表
        public static void NewExcel()
        {

            //初始化
            excel = new Excel.Application();
            books = excel.Workbooks;
            book = books.Open(@"C:\WINDOWS\Temp\属性.xlsx");    // 打开文件                         
            sheets = book.Worksheets;   // 选中文件里的所有表                      
            sheet = sheets[1];          // 选中第一个表                        
            cells = sheet.Cells;        // 选中第一个表的所有单元格      
        }
        
        //释放之前定义的对象，每循环一次就要释放一次，不然会导致Excel进程残留问题
        public static void Release(Microsoft.Office.Interop.Excel.Application excel)
        {
            //book.Close(true);           // 保存文件的修改并关闭
            excel.Quit();               // 关闭Excel
            Marshal.FinalReleaseComObject(excel);GC.Collect();
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





        /*————   ————*/

        //【向DataGridView控件粘贴数据】
        public void DataGirdViewCellPaste(DataGridView p_Data)
        {
            try
            {
                // 获取剪切板的内容，并按行分割
                string pasteText = Clipboard.GetText();
                if (string.IsNullOrEmpty(pasteText)) return;

                string[] lines = pasteText.Split(new char[] { ' ', ' ' });
                foreach (string line in lines)
                {
                    if (string.IsNullOrEmpty(line.Trim()))
                        continue;
                    // 按 Tab 分割数据
                    string[] vals = line.Split(' ');
                    p_Data.Rows.Add(vals);
                }
            }
            catch
            {
                // 不处理
            }
        }

        //下面一个问题是何时调用如上代码，可以为DataGridView创建一个右键菜单来实现，这里我展示一个通过判断键入 Ctrl+V 时来调用。代码如下：

        public void DataGridViewEnablePaste(DataGridView p_Data)
        {
            if (p_Data == null)
                return;
            p_Data.KeyDown += new KeyEventHandler(p_Data_KeyDown);
        }

        public void p_Data_KeyDown(object sender, KeyEventArgs e)
        {
            if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.V)
            {
                if (sender != null && sender.GetType() == typeof(DataGridView))
                    // 调用上面的粘贴代码
                    DataGirdViewCellPaste((DataGridView)sender);
            }
        }

    }
}


