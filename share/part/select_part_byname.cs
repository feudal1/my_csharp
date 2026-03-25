using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Text;

namespace tools
{
    public class select_part_byname
    {

        static public double run(ModelDoc2 swModel,string partName)
        {
            try
            {
 
              var asmdoc = (AssemblyDoc)swModel;
              asmdoc.GetComponentByName(partName).Select(true);
          
            }
            catch (Exception ex)
            {
                Console.WriteLine($"发生错误：{ex.Message}");
                Console.WriteLine("提示：请确保 SolidWorks 正在运行。");
            }

            return 0;
        }

    }
}