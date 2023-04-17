using SolidWorks.Interop.swdocumentmgr;
using System.Diagnostics;

namespace Sw_属性编辑工具
{
    class SwDocumentManager
    {
        public static string sLicenseKey = "IntelligentCADCAMTechnologyLtd:swdocmgr_general-11785-02051-00064-50177-08660-34307-00007-06272-33623-28314-00818-62222-42690-10633-51200-02527-59460-19873-62502-37451-32011-23810-47396-38353-45489-40357-47509-03537-04357-01293-20789-36245-47521-45501-40381-12773-37329-14337-27234-49754-58592-25690-25696-1005";

        //public const string sLicenseKey = "Axemble:swdocmgr_general-11785-02051-00064-50177-08535-34307-00007-37408-17094-12655-31529-39909-49477-26312-14336-58516-10910-42487-02022-02562-54862-24526-57604-46485-45449-00405-25144-23144-51942-23264-24676-28258-7,swdocmgr_previews-11785-02051-00064-50177-08535-34307-00007-48008-04931-27155-53105-52081-64048-22699-38918-23742-63202-30008-58372-23951-37726-23245-57604-46485-45449-00405-25144-23144-51942-23264-24676-28258-1,swdocmgr_dimxpert-11785-02051-00064-50177-08535-34307-00007-16848-46744-46507-43004-11310-13037-46891-59394-52990-24983-00932-12744-51214-03249-23667-57604-46485-45449-00405-25144-23144-51942-23264-24676-28258-8,swdocmgr_geometry-11785-02051-00064-50177-08535-34307-00007-39720-42733-27008-07782-55416-16059-24823-59395-22410-04359-65370-60348-06678-16765-23356-57604-46485-45449-00405-25144-23144-51942-23264-24676-28258-3,swdocmgr_xml-11785-02051-00064-50177-08535-34307-00007-51816-63406-17453-09481-48159-24258-10263-28674-28856-61649-06436-41925-13932-52097-22614-57604-46485-45449-00405-25144-23144-51942-23264-24676-28258-7,swdocmgr_tessellation-11785-02051-00064-50177-08535-34307-00007-13440-59803-19007-55358-48373-41599-14912-02050-07716-07769-29894-19369-42867-36378-24376-57604-46485-45449-00405-25144-23144-51942-23264-24676-28258-0";//2018版,如果正版用户，请联系代理商申请。

        //public static SwDMDocument20 swDoc { get; private set; }

        /*————  SWDocumentManager基本对象  ————*/
        public static string[] vCustPropNameArr = null;   //属性名
        public static string sCustPropStr = null;         //属性值
        public static SwDmCustomInfoType nPropType = 0;   //属性类型
        public static string txt = null;                  //属性信息输出

        //实例化swDoc对象
        public static SwDMDocument SwDMDocument(string sDocFileName)
        {
            SwDmDocumentType nDocType = 0;
            //SwDmDocumentOpenError nRetVal = 0;
            //SwDMConfigurationError results = 0;

            //SwDMApplication application = ((SwDMClassFactory)Microsoft.VisualBasic.Interaction.CreateObject("SwDocumentMgr.SwDMClassFactory.1", "")).GetApplication(sLicenseKey);
            //SwDMDocument swDoc =application.GetDocument(sDocFileName, nDocType, false, out _);


            SwDMClassFactory swClassFact = new SwDMClassFactory();
            SwDMApplication swDocMgr = swClassFact.GetApplication(sLicenseKey);
            SwDMDocument swDoc = swDocMgr.GetDocument(sDocFileName, nDocType, false, out _);
            
            
            return swDoc;
        }



        /*————  属性操作  ————*/
        //获得自定属性
        public static void CustomProperties(SwDMDocument20 swDoc)
        {
            //获得自定义属性名数组
            vCustPropNameArr = (string[])swDoc.GetCustomPropertyNames();
            //获得自定义属性值
            sCustPropStr = swDoc.GetCustomProperty(vCustPropNameArr[0], out SwDmCustomInfoType nPropType);
        }
        //删除自定属性
        public static void DelCustomProperties(SwDMDocument20 swDoc)
        {
            vCustPropNameArr = null;
            //获得自定义属性名数组
            vCustPropNameArr = (string[])swDoc.GetCustomPropertyNames(); if ((vCustPropNameArr == null)) return;
            //循环删除属性名
            foreach (string vCustPropName in vCustPropNameArr) { swDoc.DeleteCustomProperty(vCustPropName); }
            swDoc.Save();
        }
        //添加自定属性
        public static void AddCustomProperties(SwDMDocument20 swDoc, string CustPropName, string CustPropvalue)
        {
            swDoc.AddCustomProperty(CustPropName, SwDmCustomInfoType.swDmCustomInfoText, CustPropvalue);
        }
        //获得概述属性
        public static void Summary(SwDMDocument20 swDoc)
        {
            //Debug.Assert(SwDmDocumentOpenError.swDmDocumentOpenErrorNone == nRetVal);
            Debug.Print("File = " + swDoc.FullName);
            Debug.Print("");
            Debug.Print(" Version = " + swDoc.GetVersion());
            Debug.Print(" Author = " + swDoc.Author);
            Debug.Print(" Comments = " + swDoc.Comments);
            Debug.Print(" Creation Date (string) = " + swDoc.CreationDate);
            Debug.Print(" Creation Date (numeric) = " + swDoc.CreationDate2);
            Debug.Print(" Keywords = " + swDoc.Keywords);
            Debug.Print(" Last Saved By = " + swDoc.LastSavedBy);
            Debug.Print(" Last Saved Date (string) = " + swDoc.LastSavedDate);
            Debug.Print(" Last Saved Date (numeric) = " + swDoc.LastSavedDate2);
            Debug.Print(" Subject = " + swDoc.Subject);
            Debug.Print(" Title = " + swDoc.Title);
            Debug.Print(" Last Update Timestamp = " + swDoc.GetLastUpdateStamp());
            Debug.Print(" Is Detached Drawing? " + swDoc.IsDetachedDrawing());
            Debug.Print("");
        }
        //获得配置属性
        public static void ProcessConfigCustomProperties(SwDMConfiguration14 swCfg)
        {
            string[] vCustPropNameArr = null;
            string sCustPropStr = null;
            SwDmCustomInfoType nPropType = 0;

            vCustPropNameArr = (string[])swCfg.GetCustomPropertyNames();

            if ((vCustPropNameArr == null)) return;
            Debug.Print(" Configuration Custom Properties:");

            foreach (string vCustPropName in vCustPropNameArr)
            {
                sCustPropStr = swCfg.GetCustomProperty(vCustPropName, out nPropType);
                Debug.Print("   Prefaced     = " + vCustPropName + " <" + nPropType + "> = " + sCustPropStr);

                sCustPropStr = swCfg.GetCustomProperty2(vCustPropName, out nPropType);
                Debug.Print("   Not prefaced = " + vCustPropName + " <" + nPropType + "> = " + sCustPropStr);

                Debug.Print("");
            }
            Debug.Print("");
        }



        /*————  参考操作  ————*/




        /*————  其他操作  ————*/
        //文件配置信息
        public static void Configration(SwDMDocument20 swDoc)
        {
            SwDMConfigurationMgr swCfgMgr = swDoc.ConfigurationManager;
            string[] vCfgNameArr = null;

            vCfgNameArr = (string[])swCfgMgr.GetConfigurationNames();
            Debug.Print(" Number of configurations = " + swCfgMgr.GetConfigurationCount());

            foreach (string vCfgName in vCfgNameArr)
            {
                SwDMConfiguration swCfg = (SwDMConfiguration)swCfgMgr.GetConfigurationByName(vCfgName);

                Debug.Print(" Active Configuration Name = " + swCfgMgr.GetActiveConfigurationName());
                Debug.Print(" " + vCfgName);
                Debug.Print(" Description = " + swCfg.Description);
                Debug.Print(" Parent Configuration Name = " + swCfg.GetParentConfigurationName());
                Debug.Print(" Last Update Timestamp = " + swCfg.GetLastUpdateStamp());
                //Debug.Print(" 3DExperience configuration type as defined by swDm3DExperienceCfgType_e  = " + swCfg.Get3DExperienceType());
                //Debug.Print("    3DExperience physical product represented by this configuration  = " + swCfg.GetRepresentationParent());
                Debug.Print("");
            }
        }
        //文件焊件信息
        public static void CutListItems()//SolidWorks DocumentManager LicenseKey
        {
            const string sLicenseKey = "Axemble:swdocmgr_general-11785-02051-00064-50177-08535-34307-00007-37408-17094-12655-31529-39909-49477-26312-14336-58516-10910-42487-02022-02562-54862-24526-57604-46485-45449-00405-25144-23144-51942-23264-24676-28258-7,swdocmgr_previews-11785-02051-00064-50177-08535-34307-00007-48008-04931-27155-53105-52081-64048-22699-38918-23742-63202-30008-58372-23951-37726-23245-57604-46485-45449-00405-25144-23144-51942-23264-24676-28258-1,swdocmgr_dimxpert-11785-02051-00064-50177-08535-34307-00007-16848-46744-46507-43004-11310-13037-46891-59394-52990-24983-00932-12744-51214-03249-23667-57604-46485-45449-00405-25144-23144-51942-23264-24676-28258-8,swdocmgr_geometry-11785-02051-00064-50177-08535-34307-00007-39720-42733-27008-07782-55416-16059-24823-59395-22410-04359-65370-60348-06678-16765-23356-57604-46485-45449-00405-25144-23144-51942-23264-24676-28258-3,swdocmgr_xml-11785-02051-00064-50177-08535-34307-00007-51816-63406-17453-09481-48159-24258-10263-28674-28856-61649-06436-41925-13932-52097-22614-57604-46485-45449-00405-25144-23144-51942-23264-24676-28258-7,swdocmgr_tessellation-11785-02051-00064-50177-08535-34307-00007-13440-59803-19007-55358-48373-41599-14912-02050-07716-07769-29894-19369-42867-36378-24376-57604-46485-45449-00405-25144-23144-51942-23264-24676-28258-0";//2018版,如果正版用户，请联系代理商申请。

            SwDMClassFactory swClassFact = default(SwDMClassFactory);
            SwDMApplication swDocMgr = default(SwDMApplication);
            SwDMDocument10 swDocument10 = default(SwDMDocument10);
            SwDMDocument13 swDocument13 = default(SwDMDocument13);
            string sDocFileName = null;
            SwDmDocumentType nDocType = default(SwDmDocumentType);
            SwDmDocumentOpenError nRetVal = default(SwDmDocumentOpenError);


            sDocFileName = (@"C:\Users\DELL\Desktop\测试模型\Conveyor Frame_&.SLDPRT");
            // Specify your model document
            nDocType = SwDmDocumentType.swDmDocumentPart;
            swClassFact = new SwDMClassFactory();

            swDocMgr = swClassFact.GetApplication(sLicenseKey);
            swDocument10 = (SwDMDocument10)swDocMgr.GetDocument(sDocFileName, nDocType, false, out nRetVal);
            swDocument13 = (SwDMDocument13)swDocument10;

            Debug.Print("Last Update Stamp: " + swDocument13.GetLastUpdateTimeStamp());

            object[] vCutListItems = null;
            vCutListItems = (object[])swDocument13.GetCutListItems2();

            SwDMCutListItem2 Cutlist = default(SwDMCutListItem2);
            int I = 0;
            SwDmCustomInfoType nType = 0;
            string nLink = null;
            int J = 0;
            object[] vPropNames = null;

            Debug.Print("获取剪切列表项");

            for (I = 0; I <= vCutListItems.GetUpperBound(0); I++)
            {
                Cutlist = (SwDMCutListItem2)vCutListItems[I];
                Debug.Print("Name : " + Cutlist.Name);
                Debug.Print(" Quantity : " + Cutlist.Quantity);
                vPropNames = (object[])Cutlist.GetCustomPropertyNames();

                if (!((vPropNames == null)))
                {
                    Debug.Print(" GET CUSTOM PROPERTIES");
                    for (J = 0; J <= vPropNames.GetUpperBound(0); J++)
                    {
                        Debug.Print(" Property Name : " + vPropNames[J]);
                        Debug.Print(" Property Value : " + Cutlist.GetCustomPropertyValue2((string)vPropNames[J], out nType, out nLink));
                        Debug.Print(" Type : " + nType);
                        Debug.Print(" Link : " + nLink);
                    }
                }

                Debug.Print("_________________________");
            }

            Cutlist = (SwDMCutListItem2)vCutListItems[0];
            Debug.Print("ADD CUSTOM PROPERTY CALLED Testing1");
            Debug.Print(" Custom Property added? " + Cutlist.AddCustomProperty("Testing1", SwDmCustomInfoType.swDmCustomInfoText, "Verify1"));
            Debug.Print(" GET CUSTOM PROPERTIES");

            vPropNames = (object[])Cutlist.GetCustomPropertyNames();

            for (J = 0; J <= vPropNames.GetUpperBound(0); J++)
            {
                Debug.Print(" Property Name : " + vPropNames[J]);
                Debug.Print(" Property Value : " + Cutlist.GetCustomPropertyValue2((string)vPropNames[J], out nType, out nLink));
                Debug.Print(" Type : " + nType);
                Debug.Print(" Link : " + nLink);
            }

            Debug.Print("_________________________");
            Debug.Print("SET NEW CUSTOM PROPERTY VALUE FOR Testing1");
            Debug.Print(" Property Value Before Setting: " + Cutlist.GetCustomPropertyValue2("Testing1", out nType, out nLink));
            Debug.Print(" New Value Set? " + Cutlist.SetCustomProperty("Testing1", "Verify3"));
            Debug.Print(" Property Value After Setting : " + Cutlist.GetCustomPropertyValue2("Testing1", out nType, out nLink));
            Debug.Print(" GET CUSTOM PROPERTIES");

            vPropNames = (object[])Cutlist.GetCustomPropertyNames();

            for (J = 0; J <= vPropNames.GetUpperBound(0); J++)
            {
                Debug.Print(" Property Name : " + vPropNames[J]);
                Debug.Print(" Property Value : " + Cutlist.GetCustomPropertyValue2((string)vPropNames[J], out nType, out nLink));
                Debug.Print(" Type : " + nType);
                Debug.Print(" Link : " + nLink);
            }

            Debug.Print("_________________________");
            Debug.Print("DELETE CUSTOM PROPERTY Testing1");
            Debug.Print(" Delete Property Value? " + Cutlist.DeleteCustomProperty("Testing1"));

            vPropNames = (object[])Cutlist.GetCustomPropertyNames();
            if (!((vPropNames == null)))
            {
                Debug.Print(" GET CUSTOM PROPERTIES");
                for (J = 0; J <= vPropNames.GetUpperBound(0); J++)
                {
                    Debug.Print(" Property Name : " + vPropNames[J]);
                    Debug.Print(" Property Value : " + Cutlist.GetCustomPropertyValue2((string)vPropNames[J], out nType, out nLink));
                    Debug.Print(" Type : " + nType);
                    Debug.Print(" Link : " + nLink);
                }
            }

            Debug.Print("_________________________");

            //swDocument10.Save()

            swDocument10.CloseDoc();
        }


    }
}
