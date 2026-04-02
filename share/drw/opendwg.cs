using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using cad_tools;

namespace tools
{
    public class opendwg
    {
        static public void run(ModelDoc2 swModel, SldWorks swApp)
        {

            string fullpath = swModel.GetPathName();

            string? directory = Path.GetDirectoryName(fullpath);
            if (string.IsNullOrEmpty(directory))
            {
                Console.WriteLine("错误：无法获取文件所在目录。");

            }









            var drawingDoc = (DrawingDoc)swModel;

            var swSheet = (Sheet)drawingDoc.IGetCurrentSheet();
            var swViews = (object[])swSheet.GetViews();
            var partDoc = ((SolidWorks.Interop.sldworks.View)swViews[1]).ReferencedDocument;

            var thickness = get_thickness.run(partDoc);
            Debug.WriteLine($"{partDoc.GetPathName()},thickness:{thickness}");
            string outputfile = directory + "\\" + "出图" + "\\" + "工程图" + "\\" + thickness.ToString();
            string dwgFileName = outputfile + "\\" + Path.GetFileNameWithoutExtension(fullpath) + ".dwg";
            if (File.Exists(dwgFileName))
            {

                open_cad_doc_by_name.run(dwgFileName);
            }
            else
            {
                Console.WriteLine($"错误：无法找到工程图。{outputfile}");

            }

        }
    }
}