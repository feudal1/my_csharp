using System;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace tools
{
    public class drw2png
    {
        static public string run(ModelDoc2 swModel, SldWorks swApp )
        {
            
            string fullpath = swModel.GetPathName();
           
            string? directory = Path.GetDirectoryName(fullpath);
            if (string.IsNullOrEmpty(directory))
            {
                Console.WriteLine("错误：无法获取文件所在目录。");
                
            }
            
            string outputfile = directory + "\\" + "出图" + "\\" + "工程图";
            if (!Directory.Exists(outputfile))
            {
                Directory.CreateDirectory(outputfile);
            }
            string dwgFileName = directory + "\\" + "出图" + "\\" +  "工程图" + "\\" + Path.GetFileNameWithoutExtension(fullpath) + ".PNG";
            ModelView myModelView = (ModelView)swModel.ActiveView;
            myModelView.FrameState = (int)swWindowState_e.swWindowMaximized;
             swModel.ViewZoomtofit2();
            int errors=0, warnings=0;
            var result = swModel.SaveAs4(
                dwgFileName, 
                (int)swSaveAsVersion_e.swSaveAsCurrentVersion, 
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent, 
               ref errors, 
                ref warnings);
                
            Console.WriteLine($"{result}，已创建工程图{dwgFileName}");
            return dwgFileName;
        }
    }
}