using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Sw_toolkit
{
    public partial class Form4 : Form
    {

        public string[] lists = null;//列表长度
        public string[] filenames = null;

        string standard = null;

        public Form4()
        {
            InitializeComponent();
        }
        //材料文档属性标准的路径
        private void button1_Click(object sender, EventArgs e)
        {
            standard = File.Addfiles(false)[0];
            textBox1.Text = standard;
        }
        //选择文件
        private void button2_Click(object sender, EventArgs e)
        {
            filenames = File.Addfiles(true);
            for (int i = 0; i < filenames.Length; i++) { textBox2.Text += filenames[i].Trim(' ') + ";" + "\r\n"; }
            loading_Initialize(filenames.Length);
        }
        private void button3_Click(object sender, EventArgs e)
        {
            
            ISldWorks swApp = SwModelDoc.ConnectToSolidWorks(); if (swApp == null) { return; }
            if (textBox1.Text == "请选择文件" || textBox1.Text == ""|| textBox2.Text == "请选择文件" || textBox2.Text == "") { MessageBox.Show(@"请选择文件"); return; }

            //循环操作   
            for (int i = 0; i < filenames.Length; i++)
            {
                ModelDoc2 swDoc = swApp.OpenDoc(filenames[i], (int)swDocumentTypes_e.swDocPART);
                if (swDoc == null)
                {
                    swDoc = swApp.OpenDoc(filenames[i], (int)swDocumentTypes_e.swDocASSEMBLY);
                    if (swDoc == null)
                    {
                        swDoc = swApp.OpenDoc(filenames[i], (int)swDocumentTypes_e.swDocDRAWING);
                    }
                }
                swDoc.Extension.LoadDraftingStandard(standard);

                swDoc.EditRebuild3();
                swDoc.Save(); swApp.CloseDoc(filenames[i]);

                //更新进度
                loading();
            }
            MessageBox.Show(@"完成"); this.Close();

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
