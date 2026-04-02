
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

using System.Diagnostics;
namespace tools
{
    public class open_select_dwg
    {
        static public void run(ModelDoc2 swModel)
        {
           
            var swSelMgr = (SelectionMgr)swModel.SelectionManager;
            var selboj = swSelMgr.IGetSelectedObject(1);
            var face= (Face2)selboj;
            var body = (IBody2)face.GetBody();
            if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                var comp = (Component2)swSelMgr.GetSelectedObjectsComponent(1);
                swModel = (ModelDoc2)comp.GetModelDoc2();
            }

            Console.WriteLine(body.Name);
            var features=(object[])body.GetFeatures();
            var swFeature=(IFeature)(features[0]);
            
            string outputfile="";
            var fullPath = swModel.GetPathName();
            var  partname = Path.GetFileNameWithoutExtension(fullPath);

            if (string.IsNullOrEmpty(fullPath))
            {
                Console.WriteLine("错误：文档尚未保存，请先保存文件。");
                    
            }
            var directory = Path.GetDirectoryName(fullPath);
            if (string.IsNullOrEmpty(directory))
            {
                Console.WriteLine("错误：无法获取文件所在目录。");
                  
            }
            Console.WriteLine("get_select_type");
            if (swFeature.GetTypeName2() == "SheetMetal")
            {
                SheetMetalFeatureData swSheetMetalData = (SheetMetalFeatureData)swFeature.GetDefinition();
                var thickness =  Math.Round(swSheetMetalData.Thickness*1000,2).ToString();
                outputfile = directory + "\\"+"出图"+"\\" + "下料" + "\\" + thickness;
                if (!Directory.Exists(outputfile))
                {
                    Directory.CreateDirectory(outputfile);
                }
            }
 

           var dwgFileName = outputfile + "\\" +  partname+"_"+body.Name + ".dwg";
            if(File.Exists(dwgFileName))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = dwgFileName,
                    UseShellExecute = true
                });
                Console.WriteLine($"已打开：{dwgFileName}");
            }
            else
            {
                Console.WriteLine($"没有 dxf,{dwgFileName}");
            }
        }
    }
}