---
title: eDrawings_另存为Html
data: 2024-11-01 16:38:33
tag: SOLIDWORKS
toc: true
---

# eDrawings_另存为Html

```c#
 class SW_Visualize
    {
        //IVisualizeAddinManager类型转换失败，不知啥问题。。。
        public static void Render(SldWorks swApp)//导出PDF
        {
            //获得Visualize插件对象
            IVisualizeAddin testViz = swApp.GetAddInObject("{2c776b61-cdbf-4a5e-b76c-fde2df860fea}"); 
            //IVisualizeAddin testViz = (IVisualizeAddin)swApp.GetAddInObject("SolidWorks.Visualize.Implementation.VisualizeAddin.18");
            IVisualizeAddinManager visualizeAddinMgr = testViz.GetAddinManager();

            //渲染设置
            IRenderOptions vizRenderOptions = visualizeAddinMgr.RenderOptions;
            vizRenderOptions.ImageFormat = ImageFormat_e.JPEG;
            vizRenderOptions.FrameCount = 1000;
            vizRenderOptions.Width = 800;
            vizRenderOptions.Height = 800;
            vizRenderOptions.JobName = "Toaster";
            vizRenderOptions.OutputFolder = @"E:\SOLIDWORKS Visualize Content\Images";
            vizRenderOptions.DenoiserEnabled = false;

            //启动渲染
            visualizeAddinMgr.Render();
        }

    }
```



## 效



# 参考

[Save Method (IEModelViewControl) - 2022 - SOLIDWORKS API Help](https://help.solidworks.com/2022/English/api/emodelapi/eDrawings.Interop.EModelViewControl~eDrawings.Interop.EModelViewControl.IEModelViewControl~Save.html?verRedirect=1)

