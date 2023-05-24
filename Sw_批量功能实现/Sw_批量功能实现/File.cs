using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace Sw_toolkit
{
    public static class File
    {
        /*———— 获得文件对象操作  ————*/
        /// <summary>
        /// 添加文件，多选true单选false
        /// </summary>
        public static string[] Addfiles(bool isMulti)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = isMulti;
            fileDialog.Filter = "所有文件|*.*|零件|*.sldprt|装配体|*.sldasm|工程图|*.slddrw";
            fileDialog.ShowDialog();
            string[] filenames = fileDialog.FileNames;

            string[] blank_error= { "请选择文件" };
            if (filenames.Length == 0) { return blank_error; }
            return filenames;
        }
    }
}
