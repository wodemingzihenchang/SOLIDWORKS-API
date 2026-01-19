---
title: eDrawings_获得对象
data: 2024-11-01 16:38:33
tag: SOLIDWORKS
toc: true
---

# eDrawings_获得对象

```c#
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using eDrawingHostControl;
namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        eDrawingHostControl.eDrawingControl ctrl = null;
        public Form1()
        {
            InitializeComponent();
            if (null == ctrl)
            {
                ctrl = new eDrawingControl();
            }
            this.Controls.Add(ctrl);            
        }
        private void Form1_Load(object sender, System.EventArgs e)
        {
        }
        private void button1_Click(object sender, System.EventArgs e)
        {
            if(ctrl != null)
            {
            ctrl.Location = new Point(0,0);
            ctrl.Size = new System.Drawing.Size(this.Size.Width,this.Size.Height);
            ctrl.eDrawingControlWrapper.OpenDoc("C:\\Users\\Public\\Documents\\SOLIDWORKS\\SOLIDWORKS 2020\\samples\\tutorial\\EDraw\\claw\\claw.sldprt", false, false, false, "");
            }
        }
    }
}
```



## 效



# 参考

[Save Method (IEModelViewControl) - 2022 - SOLIDWORKS API Help](https://help.solidworks.com/2022/English/api/emodelapi/eDrawings.Interop.EModelViewControl~eDrawings.Interop.EModelViewControl.IEModelViewControl~Save.html?verRedirect=1)

