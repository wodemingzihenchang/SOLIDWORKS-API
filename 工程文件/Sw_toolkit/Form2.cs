using System;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;

namespace Sw_toolkit
{
    public partial class Form2 : Form
    {
        public static SldWorks swApp = (SldWorks)SwModelDoc.ConnectToSolidWorks();//获取当前运行的程序 

        public Form2()
        {
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {


        }

        private void button1_Click(object sender, EventArgs e)
        {
            string text1 = textBox1.Text;
            string text2 = textBox2.Text;
            if (text1 != null && text2 != null)
            {
                swApp.CommandInProgress = true;
                SwAssembly.Rename(text1, text2);
                swApp.CommandInProgress = false;

                MessageBox.Show("完成");
            }

        }
    }
}
