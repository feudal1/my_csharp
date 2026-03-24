using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace tools
{
    public class rename
    {
        static public void run(SldWorks swApp, ModelDoc2 swModel)
        {
            try
            {
                var partdoc = swModel;
                string fullpath= swModel .GetPathName();
                    
            var new_name = swModel.GetTitle().Replace("零件", "part");
            swModel.Extension.RenameDocument(new_name);
            
            
            }
            catch (Exception ex)
            {
                Console.WriteLine($"发生错误：{ex.Message}");
                Console.WriteLine("提示：请确保 SolidWorks 正在运行。");
            }
        }
    }
}