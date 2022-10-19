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
        /// <summary>
        /// 添加文件，多选true单选false
        /// </summary>
        public static string[] Addfiles(bool isMulti)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = isMulti;
            fileDialog.Filter = "零件|*.sldprt|装配体|*.sldasm|工程图|*.slddrw|所有文件|*.*";
            fileDialog.ShowDialog();
            string[] filenames = fileDialog.FileNames;
            return filenames;
        }

        /// <summary>
        /// 输出文件位置
        /// </summary>
        /// <returns></returns>
        public static string Outfile()   //该对话框自带的UI不太好，可改用Ookii.Dialogs，
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";
            dialog.ShowDialog();
            string Path = dialog.SelectedPath; return Path;
        }

        /// <summary>
        /// 把txt内容输入到窗口文本框
        /// </summary>
        /// <returns>txt字符串</returns>
        public static string Intxt()
        {
            StreamReader textreader = new StreamReader(File.Addfiles(false)[0]);//实例化文件流对象
            string text = ""; string textline;
            while ((textline = textreader.ReadLine()) != null) { text += textline + "\r\n"; }//循环获得txt内容
            Debug.Print(text); return text;
        }

        /// <summary>
        /// 把字符串数据输出成txt文件
        /// </summary>
        public static void Outtxt(string text)
        {
            using (FileStream fs = new FileStream(@"C:\temp\无标题.txt", FileMode.OpenOrCreate))
            {
                string txt = text;
                byte[] buffer = Encoding.Default.GetBytes(txt);
                fs.Write(buffer, 0, buffer.Length);
            }
            Process.Start(@"C:\temp\无标题.txt");
        }

        /// <summary>
        /// UTF-8字符转数字
        /// </summary>
        /// <param name="str"></param>
        public static string Utf8toint(string str)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(str);
            string[] strArr = new string[bytes.Length];

            string outstr = "";
            for (int i = 0; i < bytes.Length; i++)
            {
                strArr[i] = bytes[i].ToString("");
                outstr = outstr + strArr[i];
            }
            Console.WriteLine(outstr); return outstr;

            /*Console.WriteLine("从16进制转换回汉字：");
            for (int i = 0; i < strArr.Length; i++)
            {
                bytes[i] = byte.Parse(strArr[i], System.Globalization.NumberStyles.HexNumber);
            }
            string ret = Encoding.Unicode.GetString(bytes);
            Console.WriteLine(ret);*/
        }

        /// <summary>
        /// 利用字符用GetHashCode()方法串输出随机数0~255
        /// </summary>
        /// <returns></returns>
        public static int[] StrToint(string str)
        {
            int number = -414239725;
            //随机数方法1
            Random r = new Random();
            number = r.Next(0, 255);
            //随机数方法2
            int b = str.GetHashCode(); Debug.WriteLine(b);
            //将长字符转255输出
            int number1 = b % 1000; while (number1 < 0 || number1 > 255)
            {
                if (number1 > 255)
                {
                    number1 -= 255;
                }
                else if (number1 < 0)
                {
                    number1 += 255;
                }
            }
            int number2 = b / 1000 % 1000; while (number2 < 0 || number2 > 255)
            {
                if (number2 > 255)
                {
                    number2 -= 255;
                }
                else if (number2 < 0)
                {
                    number2 += 255;
                }
            }
            int number3 = b / 1000000 % 1000; while (number3 < 0 || number3 > 255)
            {
                if (number3 > 255)
                {
                    number3 -= 255;
                }
                else if (number3 < 0)
                {
                    number3 += 255;
                }
            }

            int[] numbers = { number1, number2, number3 };
            Debug.WriteLine(numbers[0] + "," + numbers[1] + "," + numbers[2]);
            return numbers;
        }


    }

    class FileArchives
    {
        [DllImport("kernel32.dll", EntryPoint = "SetProcessWorkingSetSize")]
        public static extern int SetProcessWorkingSetSize(IntPtr process, int minSize, int maxSize);

    }
}
