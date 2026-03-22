using System.Diagnostics;

namespace tools;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
public class asm2bom
{
    static public void run(SldWorks swApp, ModelDoc2 swModel)
    {
        var partname=Path.GetFileNameWithoutExtension(swModel.GetPathName());
        Debug.WriteLine(partname);
        
    }
}