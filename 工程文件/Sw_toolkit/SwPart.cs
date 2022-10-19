using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Diagnostics;

namespace Sw_toolkit
{
    class SwPart
    {
        public static SldWorks swApp = (SldWorks)SwModelDoc.ConnectToSolidWorks();//获取当前运行的程序
        public static ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
        public static PartDoc swPart = (PartDoc)swApp.ActiveDoc;


        /// <summary>
        /// 零件冻结全部特征
        /// </summary>
        /// <param name="OpenPart"></param>
        public static void FreezeBar(string OpenPart)
        {
            int longstatus = 0;
            swApp.ActivateDoc2(OpenPart, false, ref longstatus); //激活路径下的文件
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swUserEnableFreezeBar, true); //开启冻结栏
            swModel = (ModelDoc2)swApp.ActiveDoc;

            if (swModel.GetType() != 1) { return; }//判断为零件
            else
            {
                swModel.ForceRebuild3(false);
                swModel.FeatureManager.EditFreeze((int)swMoveFreezeBarTo_e.swMoveFreezeBarToEnd, "", true);
                swModel.Save(); swApp.CloseDoc(swModel.GetPathName());
            }
        }

        /// <summary>
        /// 面组对象获得
        /// </summary>
        public static void Getface()
        {
            object[] Bodiesobj = (object[])swPart.GetBodies(-1);
            for (int i = 0; i < Bodiesobj.Length; i++)
            {
                //获得实体对象,零件单个实体的情况Bodiesobj[0]
                Body2 componentBody = (Body2)Bodiesobj[i];

                //获得面数量
                int FacesCount = componentBody.GetFaceCount();
                Debug.WriteLine("面数量: " + FacesCount);

                //获得圆柱面数量
                int Cylinderfaces = 0;
                object[] swFaces = (object[])componentBody.GetFaces();
                for (int ii = 0; ii < swFaces.Length; ii++)
                {
                    Face2 swFace = (Face2)swFaces[ii];
                    Surface swSurFace = (Surface)swFace.GetSurface();
                    if (swSurFace.IsCylinder()) { Cylinderfaces += 1; }
                }
                Debug.WriteLine("圆柱面数量: " + Cylinderfaces);
            }
        }

        /// <summary>
        /// 遍历特征对象
        /// </summary>
        public static void TraverseFeatures(bool isTopLevel)
        {
            //获得第一个特征，并赋值到当前特征curFeat
            Feature thisFeat = (Feature)swModel.FirstFeature();
            Feature curFeat = default(Feature); curFeat = thisFeat;

            //当前特征非空就继续输出特征信息
            while ((curFeat != null))
            {
                //输出特征名称
                Debug.Print(curFeat.Name);

                //特征操作
                bool isShowDimension = false;
                if (isShowDimension == true) ShowDimensionForFeature(curFeat);


                /*//子特征的遍历
                Feature subfeat = default(Feature);
                subfeat = (Feature)curFeat.GetFirstSubFeature();
                while ((subfeat != null))
                {
                    //if (isShowDimension == true) ShowDimensionForFeature(subfeat);
                    TraverseFeatures(subfeat, false);
                    Feature nextSubFeat = default(Feature);
                    nextSubFeat = (Feature)subfeat.GetNextSubFeature();
                    subfeat = nextSubFeat; nextSubFeat = null;
                }
                subfeat = null;*/

                //进入下一个特征
                Feature nextFeat = default(Feature);
                if (isTopLevel) { nextFeat = (Feature)curFeat.GetNextFeature(); }
                else { nextFeat = null; }

                curFeat = nextFeat; nextFeat = null;
            }
        }

        /// <summary>
        /// 遍历特征中的所有尺寸
        /// </summary>
        /// <param name="feature"></param>
        public static void ShowDimensionForFeature(Feature feature)
        {
            DisplayDimension thisDisplayDim = (DisplayDimension)feature.GetFirstDisplayDimension();

            while (thisDisplayDim != null)
            {
                var dimen = (Dimension)thisDisplayDim.GetDimension();

                Debug.Print($"---特征 {feature.Name} 尺寸-->" + dimen.GetNameForSelection() + "-->" + dimen.Value + dimen.FullName);

                thisDisplayDim = (DisplayDimension)feature.GetNextDisplayDimension(thisDisplayDim);
            }
        }


        public static void Cut4()
        {

            Feature swFeature = swModel.Extension.GetLastFeatureAdded();

            swModel.Extension.SelectByID2(swFeature.Name, "SKETCH", 0, 0, 0, false, 0, null, 0);

            double depth = 0.0000000514;

            //
            swModel.FeatureManager.FeatureCut4(true, false, false, 0, 0, depth, depth, false, false, false, false, 1.74532925199433E-02, 1.74532925199433E-02, false, false, false, false, false, true, true, true, true, false, 0, 0, false, false);
            //

            MassProperty swMassProperty = swModel.Extension.CreateMassProperty();

            double swMass = swMassProperty.Mass;

            Debug.Print("Mass is : " + swMass * 1000 + " g");
        }


    }
}
