using System;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace tools
{
    public class drw2dxf
    {
        static public string run(ModelDoc2 swModel, SldWorks swApp )
        {
            
            string fullpath = swModel.GetPathName();
           
            string? directory = Path.GetDirectoryName(fullpath);
            if (string.IsNullOrEmpty(directory))
            {
                Console.WriteLine("错误：无法获取文件所在目录。");
                
            }
            
            string folderName = Path.GetFileName(directory);
            string outputRootName = string.IsNullOrWhiteSpace(folderName) ? "钣金" : $"{folderName}钣金";
            string outputfile = Path.Combine(directory, outputRootName, "工程图");
            if (!Directory.Exists(outputfile))
            {
                Directory.CreateDirectory(outputfile);
            }
            string dwgFileName = Path.Combine(outputfile, Path.GetFileNameWithoutExtension(fullpath) + ".dxf");
            swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfOutputNoScale, 1);
            swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfVersion, (int)swDxfFormat_e.swDxfFormat_R2000);
            // 设置自定义映射文件
            if (swApp != null)
            {
                string pluginDir = Path.GetDirectoryName(typeof(drw2dxf).Assembly.Location);
              
                string mapFilePath = Path.Combine(pluginDir, "dwgmaping");
              
                if (File.Exists(mapFilePath))
                {
                    swApp.SetUserPreferenceStringListValue(
                        (int)swUserPreferenceStringListValue_e.swDxfMappingFiles, 
                        mapFilePath);
                    
                    int index = swApp.GetUserPreferenceIntegerValue(
                        (int)swUserPreferenceIntegerValue_e.swDxfMappingFileIndex);
                    
                    if (index == -1)
                    {
                        var set_result=swApp.SetUserPreferenceIntegerValue(
                            (int)swUserPreferenceIntegerValue_e.swDxfMappingFileIndex, 
                            0);
                    }
                }
                else
                {
                    Console.WriteLine($"错误：无法找到自定义映射文件。{mapFilePath}");

                }
            }
            
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