using System;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
 
namespace tools
{
    public class add_type2info
    {
        static public void run(ModelDoc2 swModel)
        {
            try
            {
                       swModel.DeleteCustomInfo2("", "说明");
                       
                var thickness = get_thickness.run(swModel);
       
              
                if (thickness == 0)
                {
                    
                        bool result=swModel.AddCustomInfo2("说明", (int)swCustomInfoType_e.swCustomInfoText, "外购件");
                          Console.WriteLine($"添加说明自定义信息结果: {result}");
                }
                else
                {
                    bool result=swModel.AddCustomInfo2("说明", (int)swCustomInfoType_e.swCustomInfoText, "钣金件");
                      Console.WriteLine($"添加说明自定义信息结果: {result}");
                }
            
              
               
 
                        
                swModel.EditRebuild3();
       
 
            }
            catch (Exception ex)
            {
                Console.WriteLine($"发生错误: {ex.Message}");
                Console.WriteLine("提示：请确保 SolidWorks 正在运行。");
            }
        }
    }
}