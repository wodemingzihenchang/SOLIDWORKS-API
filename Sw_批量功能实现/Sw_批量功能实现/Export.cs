using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.IO;
using System.Windows.Forms;

namespace Sw_toolkit
{
    public partial class Export : Form
    {
        public string[] filepath = null;           //
        int Addcount = 0;                          //用于记录多次导入选择文件的计数累计（没有这项，在导入时会覆盖之前的行数据）

        //————————————//
        public Export() { InitializeComponent(); }
        private void Export_Load(object sender, EventArgs e)//许可加载
        {
            License.License_Time("2023-06-30");
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
                //进度条
                progressBar1.Value += 1;
            }
            this.dataGridView1.Rows[i + Addcount].Cells[0].Value = "/";
            Addcount += i;

            //判断是否选择文件，
            if (filepath.Length == 0) { return; }

            //进度条
            label1.Text = "进度：完成"; progressBar1.Value = progressBar1.Maximum;

        }

        private void button2_Click(object sender, EventArgs e)//运行操作
        {
            //进度条
            progressBar1.Value = 0; label1.Text = "进度： ";
            progressBar1.Maximum = dataGridView1.RowCount;

            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                //打开文件
                ModelDoc2 swDoc = SwModelDoc.OpenDoc((string)dataGridView1.Rows[i].Cells[0].Value); swDoc.EditRebuild3();
                //操作内容
                SwModelDoc.ExportToDWG2(SwModelDoc.swDoc);
                //关闭文件
                SwModelDoc.swApp.CloseDoc((string)dataGridView1.Rows[i].Cells[0].Value);
                //进度条
                this.dataGridView1.Rows[i].Cells[1].Value = "完成";
                progressBar1.Value += 1;
            }
            //进度条
            label1.Text = "进度：完成"; progressBar1.Value = progressBar1.Maximum;
            MessageBox.Show(@"完成"); this.Close();
        }

        private void button3_Click(object sender, EventArgs e)//Debug
        { }

        //————————————//
        public static string[] Addfiles(bool isMulti)//添加文件(多选true单选false)
        {
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Multiselect = isMulti,
                Filter = "零件|*.sldprt|装配体|*.sldasm|工程图|*.slddrw|所有文件| *.*"
            };
            fileDialog.ShowDialog();
            string[] filepath = fileDialog.FileNames;

            //string[] blank_error = { "请选择文件" };
            //if (filepath.Length == 0) { return blank_error; }
            return filepath;
        }
        public static void writeLogException(Exception ex) //异常信息写入日志
        {
            //获取异常信息的类、行号、异常 信息
            string exceptionStr =
            ex.StackTrace.ToString().Substring(ex.StackTrace.ToString().LastIndexOf('\\') + 1)
             + "  " + ex.Message;
            exceptionStr = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "  " + exceptionStr;
            //自己定义一个存储日志文件的位置
            string sFilePath = "C:\\Logs";
            string sFileName = DateTime.Now.ToString("yyyy-MM-dd") + ".log";
            sFileName = sFilePath + @"\\" + sFileName; //文件
            if (!Directory.Exists(sFilePath))
            {
                Directory.CreateDirectory(sFilePath);
            }
            FileStream fs;
            StreamWriter sw;
            if (System.IO.File.Exists(sFileName))
            {
                fs = new FileStream(sFileName, FileMode.Append, FileAccess.Write);
            }
            else
            {
                fs = new FileStream(sFileName, FileMode.Create, FileAccess.Write);
            }
            sw = new StreamWriter(fs);
            sw.WriteLine(exceptionStr);
            sw.Close();
            fs.Close();
        }
        /*————  进度条操作  ————*/
        public void loading_Initialize(int timelength)//初始化进度条
        {
            progressBar1.Maximum = timelength;      //设置最大长度值
            progressBar1.Value = 0;                 //设置当前值
            progressBar1.Step = 1;                  //设置没次增长多少
        }
        public void loading()                      //更新进度条
        {
            System.Threading.Thread.Sleep(100);    //暂停1秒
            progressBar1.Value += progressBar1.Step;//让进度条增加一次
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.bilibili.com/read/cv23871167");
        }
    }
}
