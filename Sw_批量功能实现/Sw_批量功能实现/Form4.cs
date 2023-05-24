using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Windows.Forms;

namespace Sw_toolkit
{
    public partial class Form4 : Form
    {
        public string[] filepath = null;           //
        int Addcount = 0;                          //用于记录多次导入选择文件的计数累计（没有这项，在导入时会覆盖之前的行数据）

        //————————————//
        public Form4() { InitializeComponent(); }

        private void button1_Click(object sender, EventArgs e)//Debug
        {
            //进度条
            progressBar1.Value = 0; label1.Text = "进度： ";
            progressBar1.Maximum = dataGridView1.RowCount;

            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                //打开文件
                ModelDoc2 swDoc = SwModelDoc.OpenDoc((string)dataGridView1.Rows[i].Cells[0].Value);

                //操作内容
                SwModelDoc.ExportToDWG2(SwModelDoc.swDoc);

                //保存关闭
                swDoc.EditRebuild3(); swDoc.Save(); SwModelDoc.swApp.CloseDoc((string)dataGridView1.Rows[i].Cells[0].Value);
                //进度条
                this.dataGridView1.Rows[i].Cells[1].Value = "完成";
                progressBar1.Value += 1;
            }   
            //进度条
            label1.Text = "进度：完成"; progressBar1.Value = progressBar1.Maximum;
            MessageBox.Show(@"完成"); this.Close();
        }

        private void button2_Click(object sender, EventArgs e)//选择文件
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

        private void button3_Click(object sender, EventArgs e)//运行操作
        {
            foreach (var item in filepath)
            {
                //打开文件
                ModelDoc2 swDoc = SwModelDoc.OpenDoc(item);
                //添加方程式
                Equation.Add();

                //保存关闭
                swDoc.EditRebuild3(); swDoc.Save(); SwModelDoc.swApp.CloseDoc(item);
                //更新进度
                loading();
            }
            MessageBox.Show(@"完成"); this.Close();
        }


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
    }
}
