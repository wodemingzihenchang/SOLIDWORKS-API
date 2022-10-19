using System;
using System.Xml;
using System.Xml.Linq;

namespace Sw_toolkit
{
    class EditXML
    {
        public void ReadXML()//读取xml文件
        {
            //将XML文件加载进来
            System.Xml.Linq.XDocument document = System.Xml.Linq.XDocument.Load("config.xml");
            //获取到XML的根元素进行操作
            System.Xml.Linq.XElement root = document.Root;
            //获取usually子节点
            System.Xml.Linq.XElement usually = root.Element("usually");
            //获取usually子节点中的元素值
            string temp = usually.Value;//此时temp = "am";
        }

        public void NewXML()//创建xml文件
        {
            //声明
        }

        public void SetXML()//对xml文件内容进行增删改查 //修改属性与新增实质是同一个方法
        {
            //节点.net没提供修改的方法本文也不做处理           
            XDocument xDoc = XDocument.Load("config.xml");
            XElement element = (XElement)xDoc.Element("usually1").Element("am");
            element.Remove();
            xDoc.Save("config.xml");

            //对节点进行查找并输出
            XmlDocument doc = new XmlDocument();
            doc.Load("orders.xml");
            XmlNodeList nodes = doc.SelectNodes("/System.Config/usually2");
            foreach (XmlNode node in nodes)
            {
                if (node.InnerText == "pm") Console.WriteLine(node.InnerText);
            }
        }

    }
}

