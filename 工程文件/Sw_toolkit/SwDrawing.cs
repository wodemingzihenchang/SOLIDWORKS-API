using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Diagnostics;
using System.Drawing;

namespace Sw_toolkit
{
    class SwDrawing
    {
        public static SldWorks swApp = (SldWorks)SwModelDoc.ConnectToSolidWorks();//获取当前运行的程序 
        public static ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
        public static DrawingDoc swDraw = (DrawingDoc)swModel;

        /// <summary>
        /// 多种格式导出
        /// </summary>
        public static void Output_multDrawing()
        {

        }

        /// <summary>
        /// 获得工程图对象（图纸、视图、零部件）
        /// </summary>
        public static void GetSheetNames()
        {
            Sheet drwSheet = (Sheet)swDraw.GetCurrentSheet();       //获取当前工程图对象

            object[] sheetNames = (object[])swDraw.GetSheetNames(); //获取当前工程图中的所有图纸名称

            object[] views = (object[])drwSheet.GetViews();         //获取当前工程图中的所有图纸视图

            foreach (object sheetName in sheetNames)
            {
                Debug.Print((String)sheetName);
            }

            foreach (View view in views)//遍历工程图零部件,输入选择视图,输出零部件名
            {
                //选择视图激活
                DrawingComponent comp = view.RootDrawingComponent; Debug.Print(comp.Name);
                //获得子件对象
                object[] childrencomps = (object[])comp.GetChildren();
                //遍历工程图零部件
                for (int i = childrencomps.GetLowerBound(0); i <= childrencomps.GetUpperBound(0); i++)
                {
                    swModel.ClearSelection2(true);
                    Debug.Print("零部件是" + ((DrawingComponent)childrencomps[i]).Name);
                }
            }

        }

        #region 图层操作
        public static void NewLayer(string Layername)//新建图层
        {
            int[] num = File.StrToint(Layername);
            //颜色定义方法,ToArgb()方法转成32进制
            Color color = Color.FromArgb(10, num[0], num[1], num[2]);
            int colorInt = color.ToArgb();
            //删除图层，以便重新新建图层（不然设置不会变）
            //LayerMgr layerMgr = (LayerMgr)swModel.GetLayerManager();
            //layerMgr.DeleteLayer(Layername);
            //新建图层(图名，说明，颜色，线型，线粗，可见，可打印)
            swDraw.CreateLayer2(Layername, "说明", (int)colorInt, (int)swLineStyles_e.swLineCONTINUOUS, (int)swLineWeights_e.swLW_NORMAL, true, true);
        }
        public static void GetLayer()//获得图层属性
        {
            //获取当前图层对象
            var swLayerMgr = (LayerMgr)swModel.GetLayerManager();
            var layCount = swLayerMgr.GetCount();
            String[] layerList = (String[])swLayerMgr.GetLayerList();

            foreach (var lay in layerList) //遍历图层
            {
                //图层类型赋值，用于后续获取图层属性
                Layer currentLayer = (Layer)swLayerMgr.GetLayer(lay);
                if (currentLayer != null)
                {
                    var currentName = currentLayer.Name;//图名
                    var currentColor = currentLayer.Color;//颜色
                    var currentDesc = currentLayer.Description;//说明
                    var currentStype = Enum.GetName(typeof(swLineStyles_e), currentLayer.Style);//线
                    var currentWidth = currentLayer.Width;//线粗

                    #region //颜色是Ref值结果转int，得到对应的RGB值
                    int refcolor = currentColor;
                    int blue = refcolor >> 16 & 255;
                    int green = refcolor >> 8 & 255;
                    int red = refcolor & 255;
                    int colorARGB = 255 << 24 | (int)red << 16 | (int)green << 8 | (int)blue;
                    Color ARGB = Color.FromArgb(colorARGB);
                    #endregion

                    Debug.Print($"图层名称：{currentName}");
                    Debug.Print($"图层颜色：R {ARGB.R},G {ARGB.G} ,B {ARGB.B}");
                    Debug.Print($"图层描述：{currentDesc}");
                    Debug.Print($"图层线型：{currentStype}");
                    Debug.Print($"-------------------------------------");
                }
            }
        }

        /// <summary>
        /// 改变零部件线型，Compname选择零部件路径，Layername添加图层名字
        /// </summary>
        /// <param name="Compname"></param>
        /// <param name="Layername"></param>
        public static void SetLayer(string Compname, string Layername)
        {
            //获得所选择的对象
            swModel.Extension.SelectByID2(Compname, "COMPONENT", 0, 0, 0, false, 0, null, 0);
            SelectionMgr swSelMgr = (SelectionMgr)swModel.SelectionManager;
            DrawingComponent swDrawComp = (DrawingComponent)swSelMgr.GetSelectedObjectsComponent4(1, 0);

            //关闭默认文档线型
            swDrawComp.UseDocumentDefaults = false;
            //线型
            swDrawComp.SetLineStyle((int)swDrawingComponentLineFontOption_e.swDrawingComponentLineFontVisible, (int)swLineStyles_e.swLineCONTINUOUS);
            //线粗
            swDrawComp.SetLineThickness((int)swDrawingComponentLineFontOption_e.swDrawingComponentLineFontVisible, (int)swLineWeights_e.swLW_CUSTOM, 0.0003);
            //图层
            swDrawComp.Layer = Layername; swDraw.ChangeComponentLayer(Layername, true); //（图层名称，是否应用所有视图）

        }
        #endregion

        #region //保存XML文件
        public void BomtoXml()//保存一个XML文件到绘图文档所在的文件夹
        {
            bool bIsFirstSheet = false; bIsFirstSheet = true;
            // 定义一个XML文件名
            string XMLname = swModel.GetPathName().Substring(0, swModel.GetPathName().Length - 6) + "xml";
            // 使用文件流写入XML
            Scripting.FileSystemObject fso = default(Scripting.FileSystemObject);
            Scripting.TextStream XMLfile = default(Scripting.TextStream);
            fso = new Scripting.FileSystemObject();//另一种方法：fso = Interaction.CreateObject("Scripting.FileSystemObject");
            XMLfile = fso.CreateTextFile(XMLname, true, true);

            XMLfile.WriteLine("<BOMS>");
            Feature swFeat = (Feature)swModel.FirstFeature();
            while ((swFeat != null))
            {
                if ("DrSheet" == swFeat.GetTypeName())
                {
                    XMLfile.WriteLine("    <SHEET>");
                    XMLfile.WriteLine("        <NAME>" + swFeat.Name + "</NAME>");
                    bIsFirstSheet = false;
                }
                if ("BomFeat" == swFeat.GetTypeName())
                {
                    //获得BOM表对象
                    BomFeature swBomFeat = (BomFeature)swFeat.GetSpecificFeature2();
                    ProcessBomFeature(swApp, swModel, swBomFeat, XMLfile);
                }
                swFeat = (Feature)swFeat.GetNextFeature();
                if ((swFeat != null))
                {
                    if ("DrSheet" == swFeat.GetTypeName() & !bIsFirstSheet)
                    {
                        XMLfile.WriteLine("    </SHEET>");
                    }
                }
            }
            XMLfile.WriteLine("    </SHEET>");
            XMLfile.WriteLine("</BOMS>");
            XMLfile.Close();
        }
        public void ProcessTableAnn(SldWorks swApp, ModelDoc2 swModel, TableAnnotation swTableAnn, Scripting.TextStream XMLfile)
        {
            int nNumRow, nNumCol = 0;
            int j, k = 0;
            int nIndex = 0, nCount = 0, nStart = 0, nEnd = 0;

            int nNumHeader = swTableAnn.GetHeaderCount();
            Debug.Assert(nNumHeader >= 1);
            int nSplitDir = swTableAnn.GetSplitInformation(ref nIndex, ref nCount, ref nStart, ref nEnd);
            if ((int)swTableSplitDirection_e.swTableSplit_None == nSplitDir)
            {
                Debug.Assert(0 == nIndex); Debug.Assert(0 == nCount); Debug.Assert(0 == nStart); Debug.Assert(0 == nEnd);
                nNumRow = swTableAnn.RowCount;
                nNumCol = swTableAnn.ColumnCount;
                nStart = nNumHeader;
                nEnd = nNumRow - 1;
            }
            else
            {
                Debug.Assert((int)swTableSplitDirection_e.swTableSplit_Horizontal == nSplitDir);
                Debug.Assert(nIndex >= 0); Debug.Assert(nCount >= 0); Debug.Assert(nStart >= 0);
                Debug.Assert(nEnd >= nStart);
                nNumCol = swTableAnn.ColumnCount;
                if (1 == nIndex) { nStart = nStart + nNumHeader; }//为表的第一部分添加头偏移量 
            }
            XMLfile.WriteLine("            <TABLE>");
            if (swTableAnn.TitleVisible)
            {
                XMLfile.WriteLine("                <TITLE>" + swTableAnn.Title + "</TITLE>");
            }
            //表头名
            string[] sHeaderText = new string[nNumCol];
            for (j = 0; j <= nNumCol - 1; j++)
            {
                sHeaderText[j] = (string)swTableAnn.GetColumnTitle2(j, true);
                // 替换XML标记的无效字符
                sHeaderText[j] = sHeaderText[j].Replace(".", "");
                sHeaderText[j] = sHeaderText[j].Replace(" ", "_");
            }
            for (j = nStart; j <= nEnd; j++)
            {
                XMLfile.WriteLine("                <ROW>");
                for (k = 0; k <= nNumCol - 1; k++)
                {
                    XMLfile.WriteLine("                    " + "<" + sHeaderText[k] + ">" + swTableAnn.get_Text2(j, k, true) + "</" + sHeaderText[k] + ">");
                }
                XMLfile.WriteLine("                </ROW>");
            }
            XMLfile.WriteLine("            </TABLE>");
        }
        public void ProcessBomFeature(SldWorks swApp, ModelDoc2 swModel, BomFeature swBomFeat, Scripting.TextStream XMLfile)
        {
            Feature swFeat = (Feature)swBomFeat.GetFeature();
            XMLfile.WriteLine("        <BOM>");
            XMLfile.WriteLine("            <NAME>" + swFeat.Name + "</NAME>");
            object[] vTableArr = (object[])swBomFeat.GetTableAnnotations();
            foreach (object vTable_loopVariable in vTableArr)
            {
                object vTable = vTable_loopVariable;
                TableAnnotation swTable = (TableAnnotation)vTable;
                ProcessTableAnn(swApp, swModel, swTable, XMLfile);
            }
            XMLfile.WriteLine("        </BOM>");

        }
        #endregion

        //获得可见实体对象
        public static void GetVisibleEntity()
        {
            View swView = (View)((SelectionMgr)swModel.SelectionManager).GetSelectedObject6(1, -1);

            DrawingComponent swDrawingComponent = swView.RootDrawingComponent;

            Component2 component = swDrawingComponent.Component; //如果是零件则为空，装配体看零部件

            Debug.Print(swDrawingComponent.Name);
            Debug.Print("Number of edges found: " + swView.GetVisibleEntityCount2(null, 1));
            Debug.Print("Number of Vertex found: " + swView.GetVisibleEntityCount2(null, 2));
            Debug.Print("Number of Face found: " + swView.GetVisibleEntityCount2(null, 3));
            Debug.Print("Number of SilhouetteEdge found: " + swView.GetVisibleEntityCount2(null, 4));
        }

        //表格对象
        public static void GetTable()
        {
            Feature f = (Feature)swDraw.FeatureByName("材料明细表1");
            BomFeature Bom = (BomFeature)f.GetSpecificFeature();
            Debug.Print(Bom.Name);
            object[] swBomAnn = (object[])Bom.GetTableAnnotations();

            //获取一般表特性的注释数据
            TableAnnotation swTableAnnotation = (TableAnnotation)swBomAnn[0];
            bool anchorAttached = swTableAnnotation.Anchored;
            Debug.Print("Table anchored        = " + anchorAttached);
            int anchorType = swTableAnnotation.AnchorType;
            Debug.Print("Anchor type           = " + anchorType);
            int nbrColumns = swTableAnnotation.ColumnCount;
            Debug.Print("Number of columns     = " + nbrColumns);
            int nbrRows = swTableAnnotation.RowCount;
            Debug.Print("Number of rows        = " + nbrRows);

        }
    }
}


