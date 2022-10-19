using Microsoft.VisualBasic;
using SolidWorks.Interop.sldworks;
using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace Sw_toolkit
{
    public partial class Form1 : Form
    {
        public static SldWorks swApp = (SldWorks)SwModelDoc.ConnectToSolidWorks();//获取当前运行的程序 
        public static ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
        public string[] filepath = null;//批量工具，接收文件路径
        public string[] progpath = null;//批量工具，接收程序路径

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)//批量工具，选择文件
        {
            //接受添加文件的路径数组，并传递给全局变量filepath
            string[] filenames = File.Addfiles(true); filepath = filenames;
            //将文件路径写到文本框内
            for (int i = 0; i < filenames.Length; i++) { textBox1.Text += filenames[i].Trim(' ') + ";" + "\r\n"; }
        }
        private void button7_Click(object sender, EventArgs e)//批量工具，选择程序
        {
            string[] pathnames = File.Addfiles(true);
            progpath = pathnames;
            for (int i = 0; i < pathnames.Length; i++)
            {
                //从文件路径截取文件名
                string filepath = pathnames[i];
                string filename = filepath.Substring(filepath.LastIndexOf("\\") + 1);
                //将文件名写到文本框内
                textBox2.Text += filename + "，";
            }
        }
        private void button2_Click(object sender, EventArgs e)//批量工具，启动按钮
        {
            //获得字符内容，组成字符数组
            string[] filenames = filepath;
            string[] prognames = progpath;
            //初始化进度条
            progressBar1.Maximum = filenames.Length;    //设置最大长度值
            progressBar1.Value = 0;                     //设置当前值
            progressBar1.Step = 1;                      //设置没次增长多少
            //进行批量操作,循环文件数       
            for (int i = 0; i < filenames.Length; i++)
            {
                //打开textbox1的文件
                SwModelDoc.Open(filenames[i]);
                //执行textbox2的程序
                for (int ii = 0; ii < prognames.Length; ii++)
                {
                    swApp.RunMacro(prognames[ii], "", "main");//选择宏文件路径，模块名默认，程序名默认Main
                }
                //关闭模型文件                
                swApp.CloseDoc(filenames[i]);
                //更新进度条
                System.Threading.Thread.Sleep(1000);//暂停1秒
                progressBar1.Value += progressBar1.Step; //让进度条增加一次
            }
            //提示完成
        }

        private void button3_Click(object sender, EventArgs e)//输出属性
        {
            if (textBox1.Text == "") EditExcel.GetPropertiesToExcel();
            else
            {
                loading_Initialize(filepath.Length);//初始化进度条
                EditExcel.NewExcel();
                EditExcel.GetPropnameToExcel(filepath[0], 1, 1);
                for (int i = 0; i < filepath.Length; i++)
                {
                    EditExcel.GetPropvalueToExcel(filepath[i], 1, i + 2); loading();
                }

                EditExcel.Release();
                MessageBox.Show("完成");
                Process.Start(@"C:\temp\属性.xlsx");
            }
        }
        private void button6_Click(object sender, EventArgs e)//添加属性
        {
            this.Hide(); EditExcel.SetPropertiesToExcel();
        }
        private void button8_Click(object sender, EventArgs e)//删除属性
        {
            String PropNames = Interaction.InputBox("键入属性名1，属性名2...逗号分隔多个属性（默认空白是删除全部属性）", "输入框", "", -1, -1);

            if (textBox1.Text == "")
            {
                if (PropNames == "") { SwModelDoc.DelProperties(""); }
                else { SwModelDoc.DelTheProperties(PropNames); }
            }
            else
            {
                loading_Initialize(filepath.Length);//初始化进度条
                for (int i = 0; i < filepath.Length; i++)
                {
                    swApp.OpenDoc(filepath[i], 1);
                    if (PropNames == "") { SwModelDoc.DelProperties(filepath[i]); }
                    else { SwModelDoc.DelTheProperties(PropNames); }
                    loading();
                }
                MessageBox.Show("完成");
            }
            this.Close();
        }

        private void button4_Click(object sender, EventArgs e)//装配体分配多配置
        {
            //获得多配置数组
            string[] configNames = (string[])swModel.GetConfigurationNames();
            //输入需修改的表名
            String bomname = Interaction.InputBox("键入分配多配置表格名字", "输入框", "材料明细表1", -1, -1);
            //循环设置每个配置对应的表格配置
            for (int i = 0; i < configNames.Length; i++)
            {
                //激活配置                
                swModel.ShowConfiguration2(configNames[i]);
                //设置表格配置
                SwAssembly.SetBOM(bomname, configNames[i]);
            }
        }
        private void button9_Click(object sender, EventArgs e)//冻结零部件
        {
            string[] comppath = SwAssembly.GetComps(swModel.GetPathName());
            loading_Initialize(comppath.Length);//初始化进度条
            for (int i = 0; i < comppath.Length; i++) { SwPart.FreezeBar(comppath[i]); loading(); }
        }
        private void button10_Click(object sender, EventArgs e)//装配体改名
        {
            Form2 f2 = new Form2();
            f2.Show();
        }

        private void button5_Click(object sender, EventArgs e)//工程图输出多格式
        {

        }
        private void button11_Click(object sender, EventArgs e)//工程图Bom输出XML
        {
            SwDrawing swd = new SwDrawing();
            swd.BomtoXml(); MessageBox.Show("完成");
        }
        private void button12_Click(object sender, EventArgs e)//工程图线条改颜色分层
        {
            DrawingDoc swDraw = (DrawingDoc)swModel;
            Sheet drwSheet = (Sheet)swDraw.GetCurrentSheet();       //获取当前工程图对象
            object[] views = (object[])drwSheet.GetViews();
            SolidWorks.Interop.sldworks.View view = (SolidWorks.Interop.sldworks.View)views[0];

            DrawingComponent swDrawComp0 = view.RootDrawingComponent; //获取当前工程图总装配体对象
            object[] childrencomps = (object[])swDrawComp0.GetChildren();//获取当前工程图子装配体对象

            loading_Initialize(childrencomps.Length);
            for (int i = 0; i < childrencomps.Length; i++)//遍历工程图零部件
            {
                DrawingComponent swDrawComp = (DrawingComponent)childrencomps[i];
                Component2 swComp = (Component2)swDrawComp.Component;
                Debug.Print("零部件是" + swComp.Name);
                string samnename = swComp.Name.Substring(0, swComp.Name.LastIndexOf('-'));//统一同名零件
                //新建图层
                SwDrawing.NewLayer(samnename);
                //选择路径
                string selectname = swDrawComp0.Name + "@" + view.Name + "/" + swComp.Name;
                //设置图层
                SwDrawing.SetLayer(selectname, samnename); loading();
            }
            MessageBox.Show("完成");
        }



        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("https://space.bilibili.com/12254884?spm_id_from=333.1007.0.0");
        }//联系我B站
        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("https://wodemingzihenchang.github.io/");
        }//帮助文档_博客
        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("https://github.com/wodemingzihenchang/SOLIDWORKS-API");
        }//联系我Github



        public void loading_Initialize(int timelength)//初始化进度条
        {
            progressBar1.Maximum = timelength;      //设置最大长度值
            progressBar1.Value = 0;                 //设置当前值
            progressBar1.Step = 1;                  //设置没次增长多少
        }
        public void loading()
        {
            System.Threading.Thread.Sleep(100);    //暂停1秒
            progressBar1.Value += progressBar1.Step;//让进度条增加一次
        }

        public void test()
        {
            //加速读取 swApp.CommandInProgress = true;

            //SwPart.Getface();
            //SwAssembly.Rename("V1","V2" );
            //SwDrawing.GetVisibleEntity();

        }

    }
}
