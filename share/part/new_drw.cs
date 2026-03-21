using System;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace tools
{
    public class New_drw
    {
        static public void run(SldWorks swApp, ModelDoc2 swModel)
        {
            try
            {
                string fullpath= swModel .GetPathName();
                      string drwpath = swModel.GetPathName().Replace("prt", "PRT").Replace("PRT", "drw");
                 
                 // 检查工程图文件是否已存在
                 if (File.Exists(drwpath))
                 {
                     Console.WriteLine($"工程图已存在，直接打开：{drwpath}");
                  swApp.OpenDoc(drwpath, (int)swDocumentTypes_e.swDocDRAWING);
                    
                  return;
                 }
                 
              swApp.NewDocument(@"C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2023\templates\gb_a4.drwdot", 0, 0, 0);
               
               swModel = (ModelDoc2)swApp.ActiveDoc;
                  if (swModel.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
                {
                    Console.WriteLine("错误：无法创建工程图");
                    return;
                }
                else
                {
                 
                    
                        DrawingDoc drawingDoc = (DrawingDoc)swModel;
                      
                  
                drawingDoc.GenerateViewPaletteViews(fullpath);
              
               var view1 = drawingDoc.CreateDrawViewFromModelView3(fullpath, "*上视", 0.08, 0.10, 0);
                if (view1 == null) Console.WriteLine("view=null");
                    if (view1==null) view1 = drawingDoc.CreateDrawViewFromModelView3(fullpath, "*Top", 0.13, 0.22, 0);
                var view2 = drawingDoc.CreateUnfoldedViewAt3(0.20, 0.10, 0, false);
              swModel.Extension.SelectByID2(view1.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                var view3 = drawingDoc.CreateUnfoldedViewAt3(0.08, 0.15, 0, false);
               swModel.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDisplaySketches, false);

          
               swModel.EditRebuild3();
               swModel.SaveAs3(drwpath, 0, 0);
  Console.WriteLine($"成功，已创建工程图{drwpath}");
                }
            
            
            }
            catch (Exception ex)
            {
                Console.WriteLine($"发生错误：{ex.Message}");
                Console.WriteLine("提示：请确保 SolidWorks 正在运行。");
            }
        }
    }
}