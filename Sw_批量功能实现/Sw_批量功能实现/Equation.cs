using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;

namespace Sw_toolkit
{
    class Equation
    {
        //
        public static EquationMgr equationMgr = SwModelDoc.swDoc.GetEquationMgr();//方程式管理器对象


        //修改方程式
        public static void Add()
        {
            if (equationMgr != null)
            {
                equationMgr = SwModelDoc.swDoc.GetEquationMgr();
                //全局变量
                equationMgr.Add(0, "\"A\" = \"图号代码\"");
                equationMgr.Add(-1, "\"B\" = \"名称代码\"");
                //特征
                //equationMgr.Add(-1, "\"切除 - 拉伸1\" = \"suppressed\"");
                //方程式
                //equationMgr.Add(-1, "\"D1@草图1\" = 0.05in+0.5in");
            }
        }

        /*public static void Add2()
        {
            bool longEquation;

            if (equationMgr == null) Console.WriteLine("无方程式");

            //Add a global variable assignment at index, 0, to all configurations
            //在所有配置中添加索引0处的全局变量赋值
            longEquation = equationMgr.Add3(0, "\"A\" = 2in", true, (int)swInConfigurationOpts_e.swAllConfiguration, null);
            if (longEquation != false)
                ErrorMsg(swApp, "Failed to add a global variable assignment");

            //Add a dimension equation at index, 1, to all configurations
            //向所有配置添加索引1处的维度方程
            longEquation = equationMgr.Add3(1, "\"D1@Boss-Extrude1\" = 0.05in", true, (int)swInConfigurationOpts_e.swAllConfiguration, null);
            if (longEquation != true)
                ErrorMsg(swApp, "Failed to add a dimension equation");

            //Modify dimension equation at index, 1, in all configurations
            //修改所有配置中索引1处的维度方程
            longEquation = equationMgr.SetEquationAndConfigurationOption(1, "\"D1@Boss-Extrude1\" = 0.07in", (int)swInConfigurationOpts_e.swAllConfiguration, null);
            if (longEquation != true)
                ErrorMsg(swApp, "Failed to modify a dimension equation");


        }*/


        public void ErrorMsg(SldWorks swApp, string Message)
        {
            swApp.SendMsgToUser2(Message, 0, 0);
            swApp.RecordLine("'*** WARNING - General");
            swApp.RecordLine("'*** " + Message);
            swApp.RecordLine("");
        }


        //方程式输入TXT文件
        public static void Simple()
        {
            //equationMgr.FilePath = @"E:\20-LeonLin\[06]Learn\#测试模型\2022\其他模型\equations.txt";
            //equationMgr.LinkToFile = false;
            equationMgr.Add(-1, "图号代码");
        }


        /*
 Property	AngularEquationUnits	Gets or sets the angular units used in equations.  
 Property	AutomaticRebuild	Gets or sets whether to automatically rebuild after modifications.  
 Property	AutomaticSolveOrder	Gets or sets whether to automatically sequence equations in an order determined by SOLIDWORKS to produce accurate results.  
 Property	Disabled	Gets or sets whether to disable the specified equation in the model.  
 Property	Equation	Gets or sets the equation at the specified index.  
 Property	FilePath	Gets or sets the path for an exported equation text (.txt) file.  
 Property	GlobalVariable	Gets whether the equation at the specified index is a global variable.  
 Property	LinkToFile	Gets or sets whether the equation is linked to an exported equation text (.txt) file.  
 Property	Status	Gets the status of the last equation that was executed.  
 Property	Suppression	Obsolete as of SOLIDWORKS 2014 and later.  
 Property	Value	Gets the value of the equation at the specified index.   

 Method	Add	Obsolete. Superseded by IEquationMgr::Add2.  
 Method	Add2	Adds an equation at the specified index.  
 Method	Add3	Adds an equation at the specified index for the specified configurations.  
 Method	ChangeSuppressionForAllConfigurations	Changes the suppression state of the specified equation in all configurations.  
 Method	ChangeSuppressionForConfiguration	Changes the suppression state of an equation in the specified configuration.  
 Method	Delete	Deletes the equation at the specified index.  
 Method	EvaluateAll	Evaluates all equations.  
 Method	GetConfigurationOption	Gets the configuration option for the equation at the specified index.  
 Method	GetCount	Gets the number of equations in the model.  
 Method	GetDisabledEquationCount	Gets the number of disabled equations in the model.  
 Method	IAdd3	Adds an equation at the specified index for the specified configurations.  
 Method	ISetEquationAndConfigurationOption	Modifies the equation at the specified index for the specified configurations.  
 Method	SetEquationAndConfigurationOption	Modifies the equation at the specified index for the specified configurations.  
 Method	UpdateValuesFromExternalEquationFile	Updates equations dependent on a linked equation file and ensures that the linked equation file exists and updates its current path, if necessary.  


         */
    }
}
