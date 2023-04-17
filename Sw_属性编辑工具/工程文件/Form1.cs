using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using SolidWorks.Interop.swdocumentmgr;

namespace Sw_属性编辑工具
{
    public partial class Form1 : Form
    {
        public string[] filepath = null;           //
        int Addcount = 0;                          //用于记录多次导入选择文件的计数累计（没有这项，在导入时会覆盖之前的行数据）

        //————————————//
        ISwDMDocument swDoc;                       //SW文件
        public string[] vCustPropNameArr = null;   //属性名
        public string sCustPropStr = null;         //属性值
        public string[] arr22 = { "选择属性操作" };




        //——————属性操作——————//
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)//选择文件
        {
            //选择文件
            filepath = Addfiles(true);
            //进度条
            progressBar1.Value = 0; label1.Text = "进度： ";
            progressBar1.Maximum = filepath.Length;
            //设置表格行列
            dataGridView1.RowCount = filepath.Length + Addcount + 1;
            dataGridView1.ColumnCount = 2;
            dataGridView1.Columns[0].HeaderText = "文件位置"; int i;

            for (i = 0; i < filepath.Length; i++)
            {//写入表格内容
                this.dataGridView1.Rows[i + Addcount].Cells[0].Value = filepath[i];

            }
            this.dataGridView1.Rows[i + Addcount].Cells[0].Value = "/";
            Addcount += i;


            //判断是否选择文件，
            if (filepath.Length == 0) { return; }
            progressBar1.Value = 0;
            for (int j = 0; j < filepath.Length; j++)
            {//获得自定义属性名数组（并做数组concat合并）
                swDoc = SwDocumentManager.SwDMDocument((string)dataGridView1.Rows[j].Cells[0].Value);
                if ((string[])swDoc.GetCustomPropertyNames() != null)
                {
                    arr22 = arr22.Concat((string[])swDoc.GetCustomPropertyNames()).ToArray(); //提示问题：XML缺少主元素？,原因：当无属性名时，最该空属性名编辑时会提示
                }
                swDoc.CloseDoc();
                //进度条
                progressBar1.Value += 1;
            }
            //排除重复属性名
            vCustPropNameArr = arr22.Distinct().ToArray();
            //属性名填入列表，可供选择指定属性进行操作（避免全部属性都加载的情况）
            checkedListBox1.Items.Clear();
            checkedListBox1.Items.AddRange(vCustPropNameArr);
            //进度条
            label1.Text = "进度：完成"; progressBar1.Value = progressBar1.Maximum;

        }

        private void button2_Click(object sender, EventArgs e)//获得属性
        {
            //进度条
            progressBar1.Value = 0; label1.Text = "进度： ";
            progressBar1.Maximum = dataGridView1.RowCount;

            //填入属性名
            string s = "";
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {//判断是否选择指定的属性进行操作
                if (checkedListBox1.GetItemChecked(i))
                {
                    foreach (object itemChecked in checkedListBox1.CheckedItems) { s += ',' + itemChecked.ToString(); }
                    vCustPropNameArr = s.Split(',');
                    break;
                }
            }
            if (vCustPropNameArr != null)
            {//判断是否为空白属性 + 添加列数量，来接受属性名表头
                dataGridView1.ColumnCount = vCustPropNameArr.Length;
                for (int ii = 1; ii < vCustPropNameArr.Length; ii++)
                {//填入属性名列表头
                    dataGridView1.Columns[ii].HeaderText = vCustPropNameArr[ii];
                }
            }

            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {//填入属性值
                swDoc = SwDocumentManager.SwDMDocument((string)dataGridView1.Rows[i].Cells[0].Value);
                for (int iii = 1; iii < vCustPropNameArr.Length; iii++)
                {//添加列内容,获得自定义属性值
                    try
                    {//判断空值
                        sCustPropStr = swDoc.GetCustomProperty(vCustPropNameArr[iii], out SwDmCustomInfoType nPropType);
                        dataGridView1.Rows[i].Cells[iii].Value = sCustPropStr;
                    }
                    catch (Exception)
                    {//无属性留空值
                        if (sCustPropStr == null) { dataGridView1.Rows[i].Cells[iii].Value = ""; }// throw;
                    }
                }
                swDoc.CloseDoc();
                //进度条
                progressBar1.Value += 1;
            }
            //进度条
            label1.Text = "进度：完成"; progressBar1.Value = progressBar1.Maximum;
        }

        private void button3_Click(object sender, EventArgs e)//修改属性
        {
            //进度条
            progressBar1.Value = 0; label1.Text = "进度： ";
            progressBar1.Maximum = dataGridView1.RowCount;
            //定义属性名，属性值
            string CustPropName;
            string CustPropvalue;
            //获得行列数据
            DataGridViewRowCollection rows = dataGridView1.Rows;
            DataGridViewColumnCollection columns = dataGridView1.Columns;
            for (int i = 0; i < rows.Count - 1; i++)
            {//读取文件
                swDoc = SwDocumentManager.SwDMDocument((string)dataGridView1.Rows[i].Cells[0].Value); string[] ss = swDoc.GetCustomPropertyNames();
                for (int ii = 1; ii < columns.Count; ii++)
                {//读取属性                 
                    CustPropName = (string)dataGridView1.Columns[ii].HeaderText;
                    CustPropvalue = null;
                    var v = dataGridView1.Rows[i].Cells[ii].Value;
                    if (v == null) { }//判断文件属性值是否为空
                    else CustPropvalue = (string)dataGridView1.Rows[i].Cells[ii].Value.ToString();                                      

                    if (ss != null)
                    {
                        if (ss.Contains(CustPropName))
                        {//判断文件属性名是否包含表格属性名，从而决定是否修改属性值
                            swDoc.SetCustomProperty(CustPropName, CustPropvalue);
                        }
                        else
                        {//判断文件属性名不在表格属性名内的，是否添加
                            if (CustPropvalue != "") { swDoc.AddCustomProperty(CustPropName, SwDmCustomInfoType.swDmCustomInfoText, CustPropvalue); }
                            //新增的空白属性值是否填入
                            //else swDoc.AddCustomProperty(CustPropName, SwDmCustomInfoType.swDmCustomInfoText, "");
                        }
                    }
                    else
                    {//判断文件属性名不在表格属性名内的，是否添加
                        if (CustPropvalue != "") { swDoc.AddCustomProperty(CustPropName, SwDmCustomInfoType.swDmCustomInfoText, CustPropvalue); }
                        //新增的空白属性值是否填入
                        //else swDoc.AddCustomProperty(CustPropName, SwDmCustomInfoType.swDmCustomInfoText, "");
                    }

                }
                swDoc.Save(); swDoc.CloseDoc();
                progressBar1.Value += 1;//让进度条增加一次
            }
            //进度条
            label1.Text = "进度：完成"; progressBar1.Value = progressBar1.Maximum;
            MessageBox.Show("完成");
        }

        private void button6_Click(object sender, EventArgs e)//删除属性
        {
            //进度条
            progressBar1.Value = 0; label1.Text = "进度： ";
            progressBar1.Maximum = dataGridView1.RowCount;

            //填入属性名
            string s = "";
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {//判断是否选择指定的属性进行操作
                if (checkedListBox1.GetItemChecked(i))
                {
                    foreach (object itemChecked in checkedListBox1.CheckedItems) { s += ',' + itemChecked.ToString(); }
                    vCustPropNameArr = s.Split(',');
                    break;
                }
            }

            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {//删除属性值
                swDoc = SwDocumentManager.SwDMDocument((string)dataGridView1.Rows[i].Cells[0].Value);
                if (vCustPropNameArr != null)
                {
                    for (int iii = 1; iii < vCustPropNameArr.Length; iii++)
                    {//根据属性名删除对应属性
                        try
                        {//判断空值
                            swDoc.DeleteCustomProperty(vCustPropNameArr[iii]);
                        }
                        catch (Exception)
                        {//无属性留空值
                         // throw;
                        }
                    }
                }
                swDoc.Save(); swDoc.CloseDoc();

                //进度条
                progressBar1.Value += 1;
            }
            //进度条
            label1.Text = "进度：完成"; progressBar1.Value = progressBar1.Maximum;
        }


        //——————★Excel操作——————//
        private void button4_Click(object sender, EventArgs e)//保存Excel数据
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel|*.xls|Excel|*.xlsx"
            };
            if (saveFileDialog.ShowDialog() == DialogResult.OK && saveFileDialog.FileName != "")
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                if (xlApp == null) { MessageBox.Show("无Excel"); return; }

                Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
                Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
                Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.Worksheets[1];

                for (int i = 0; i < dataGridView1.ColumnCount; i++)
                {
                    worksheet.Cells[1, i + 1] = dataGridView1.Columns[i].HeaderText;
                }
                //写入数值
                for (int r = 0; r < dataGridView1.RowCount; r++)
                {
                    progressBar1.Value = dataGridView1.RowCount;
                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    {
                        worksheet.Cells[r + 2, i + 1] = dataGridView1.Rows[r].Cells[i].Value;
                    }
                }
                worksheet.Columns.EntireColumn.AutoFit();
                MessageBox.Show("资料保存成功");
                workbook.Saved = true;
                workbook.SaveCopyAs(saveFileDialog.FileName);
                xlApp.Quit();

                //进度条
                label1.Text = "进度：完成"; progressBar1.Value = progressBar1.Maximum;
            }
        }

        private void button5_Click(object sender, EventArgs e)//打开Excel数据
        {
            OpenFileDialog fd = new OpenFileDialog
            {
                Filter = "Excel(*.xls)|*.xls|Excel(*.xlsx)|*.xlsx"
            };
            string strpath;
            if (fd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    strpath = fd.FileName;

                    //string strCon = "provider=microsoft.jet.oledb.4.0;data source=" + strpath + ";extended properties=excel 8.0";//32位平台
                    //string strCon = "provider=microsoft.ACE.OLEDB.12.0" + strpath + ";extended properties=excel 8.0";//需要安装Access Data engin数据库引擎
                    string strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strpath + ";Extended Properties=\"Excel 12.0;HDR=YES\"";//64位平台可用

                    OleDbConnection Con = new OleDbConnection(strCon);
                    string strSql = "select* from [Sheet1$]";
                    OleDbCommand cmd = new OleDbCommand(strSql, Con);
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + "(需使用Excel 2003.xls版本)");
                }
            }
        }


        //————————————//
        public static string[] Addfiles(bool isMulti)//添加文件(多选true单选false)
        {
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Multiselect = isMulti,
                Filter = "零件|*.sldprt|装配体|*.sldasm|工程图|*.slddrw"
            };
            fileDialog.ShowDialog();
            string[] filepath = fileDialog.FileNames;

            //string[] blank_error = { "请选择文件" };
            //if (filepath.Length == 0) { return blank_error; }
            return filepath;
        }

        private void Debug_Click(object sender, EventArgs e)//Debug
        {
            swDoc = SwDocumentManager.SwDMDocument(@"C:\Users\DELL\Desktop\测试模型\Part1 2023.SLDPRT");
            Console.WriteLine(swDoc.GetVersion());
        }










        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/wodemingzihenchang/SOLIDWORKS-API");
        }


    }
}
