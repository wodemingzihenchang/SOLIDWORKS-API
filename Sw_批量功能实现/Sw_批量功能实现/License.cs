using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Win32;

namespace Sw_toolkit
{
    class License
    {
        //——————时间许可——————//

        public static void CreatLicense(string text)    //新建许可文件
        {
            //清空内容
            FileStream fs = new FileStream(@"C:\WINDOWS\Temp\CreatLicense.txt", FileMode.OpenOrCreate);
            fs.Seek(0, SeekOrigin.Begin);
            fs.SetLength(0);
            fs.Close();

            //添加内容
            StreamWriter sw = new StreamWriter(@"C:\WINDOWS\Temp\CreatLicense.txt", true, System.Text.Encoding.GetEncoding("gb2312"));
            sw.WriteLine(text);
            sw.Flush();
            sw.Close();
        }

        public static string GetLicense()               //读取许可文件
        {
            string license = DateTime.Now.ToString("yyyy-MM-dd").ToString();
            try
            {
                StreamReader sw = new StreamReader(@"C:\WINDOWS\Temp\CreatLicense.txt", System.Text.Encoding.UTF8);
                license = sw.ReadToEnd();
                sw.Close(); return license;
            }
            catch (Exception)
            {
                License.CreatLicense(DateTime.Now.ToString("yyyy-MM-dd").ToString()); return license;
                //throw;
            }
        }

        public static void License_Time(string deadline)//设定截止期限
        {
            //TimeSpan License_Time1 = DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")) - DateTime.Parse(GetLicense().Substring(0, 10));
            TimeSpan License_Time2 = DateTime.Parse(deadline) - DateTime.Parse(GetLicense().Substring(0, 10));

            if (License_Time2.Days < 0)
            {
                MessageBox.Show("程序更新，请联系https://space.bilibili.com/12254884", "注意", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
                Environment.Exit(0);
            }
            //程序名称显示试用倒计时提醒
            //this.Text = string.Format("Sw_属性编辑工具（您是试用天数还剩{0}天）", remainingDays);
        }


        //——————注册表许可——————//

        public static string Registry_Value;
        public static void RegistryKey()                //获得注册表信息
        {
            RegistryKey regkey = Registry.LocalMachine;
            RegistryKey sys = regkey.OpenSubKey(@"SOFTWARE\SolidWorks\Licenses\Serial Numbers\");
            Registry_Value = sys.GetValue(@"SolidWorks").ToString();

            //使用两个foreach 语句检索 HKEY_LOCAL_MACHINE\SOFTWARE 键下的所有子项目
            /*foreach (string str in sys.GetValueNames())
            {
                Console.WriteLine("子项名:" + str); RegistryKey sikey = sys.OpenSubKey(str);
                //打开子键
                //foreach (string sVName in sikey.GetValueNames()) { Console.WriteLine(sVName + sikey.GetValue(sVName)); }
            }*/
        }

        public static void License_Registry()           //注册表内容许可
        {
            RegistryKey();
            if (Registry_Value.StartsWith("9000") || Registry_Value.StartsWith("9010")) { }
            else { MessageBox.Show("程序更新，请联系https://space.bilibili.com/12254884", "注意", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk); Environment.Exit(0); }
        }

    }
}
