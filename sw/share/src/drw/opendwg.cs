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
        static public void run(ModelDoc2 swModel, SldWorks swApp,bool opencaxa)
        {

            string fullpath = swModel.GetPathName();

            string? directory = Path.GetDirectoryName(fullpath);
            if (string.IsNullOrEmpty(directory))
            {
                Console.WriteLine("错误：无法获取文件所在目录。");
                return;
            }




            var drawingDoc = (DrawingDoc)swModel;

            var swSheet = (Sheet)drawingDoc.IGetCurrentSheet();
            var swViews = (object[])swSheet.GetViews();
            var partDoc = ((SolidWorks.Interop.sldworks.View)swViews[1]).ReferencedDocument;

            string outputRoot = ResolveOutputRoot(directory);
            string outputfile;
            bool is_cnc = false;
            
            if (partDoc.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                Debug.WriteLine($"{partDoc.GetPathName()},type:assembly");
                outputfile = Path.Combine(outputRoot, "焊接图");
            }
            else
            {
                var thickness = get_thickness.run(partDoc);
                Debug.WriteLine($"{partDoc.GetPathName()},thickness:{thickness}");

                if (thickness == 0)
                {
                    outputfile = Path.Combine(outputRoot, "CNC");
                    is_cnc = true;
                }
                else
                {
                    outputfile = Path.Combine(outputRoot, "工程图", thickness.ToString());
                }
            }

            string dwgFileName = Path.Combine(outputfile, Path.GetFileNameWithoutExtension(fullpath) + ".dwg");
            if (File.Exists(dwgFileName))
            {
                if (opencaxa)
              
                {
                    open_cad_doc_by_shell.run(dwgFileName);
                }
                else
                {
                   
                    open_cad_doc_by_name.run( dwgFileName);
                }
            }
            else
            {
                Console.WriteLine($"错误：无法找到工程图。{dwgFileName}");

            }

        }

        private static string ResolveOutputRoot(string modelDirectory)
        {
            string folderName = Path.GetFileName(modelDirectory);
            string preferredRootName = string.IsNullOrWhiteSpace(folderName) ? "钣金" : $"{folderName}钣金";
            return Path.Combine(modelDirectory, preferredRootName);
        }
    }
}