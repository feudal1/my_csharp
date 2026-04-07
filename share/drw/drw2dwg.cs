using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace tools
{
    public class drw2dwg
    {
        static public string run(ModelDoc2 swModel, SldWorks swApp)
        {
            
            string fullpath = swModel.GetPathName();
           
            string? directory = Path.GetDirectoryName(fullpath);
            if (string.IsNullOrEmpty(directory))
            {
                Console.WriteLine("错误：无法获取文件所在目录。");
                
            }
            
          
         
            Debug.WriteLine($"正在转换Drw为DWG。{fullpath}");
            // 设置自定义映射文件
            if (swApp != null)
            {
                string pluginDir = Path.GetDirectoryName(typeof(drw2dwg).Assembly.Location);
              
                string mapFilePath = Path.Combine(@"C:\Users\Administrator\", "dwgmaping");
              
                if (File.Exists(mapFilePath))
                {
                    Debug.WriteLine($"已设置自定义映射文件。{mapFilePath}");
                    swApp.SetUserPreferenceStringListValue(
                        (int)swUserPreferenceStringListValue_e.swDxfMappingFiles, 
                        mapFilePath);
                    
                  
                }
                else
                {
                    Debug.WriteLine($"错误：无法找到自定义映射文件。{mapFilePath}");

                }
            }

    

          
           var drawingDoc = (DrawingDoc)swModel;
          
            var swSheet = (Sheet)drawingDoc.IGetCurrentSheet();
            var swViews = (object[])swSheet.GetViews();
            var partDoc = ((SolidWorks.Interop.sldworks.View)swViews[1]).ReferencedDocument;
            
            swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfOutputNoScale, 1); 
            
            string outputfile;
            
            // 先判断是否为装配体，是装配体就不是 CNC
            if (partDoc.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                Debug.WriteLine($"{partDoc.GetPathName()},type:assembly");
                swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfOutputFonts, 1);
                swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfVersion, (int)swDxfFormat_e.swDxfFormat_R2000);
                outputfile = directory + "\\" + "出图" + "\\" + "焊接图";
            }
            else
            {
                var thickness = get_thickness.run(partDoc);
                Debug.WriteLine($"{partDoc.GetPathName()},thickness:{thickness}");
                
                if (thickness == 0)
                {
                    swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfVersion, (int)swDxfFormat_e.swDxfFormat_R12);
                    outputfile = directory + "\\" + "出图" + "\\" + "CNC";
                }
                else
                {
                    swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfOutputFonts, 1);
                    swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfVersion, (int)swDxfFormat_e.swDxfFormat_R2000);
                    var meterialDB = "";
                    string meterial=((PartDoc)partDoc).GetMaterialPropertyName( out meterialDB);
                    
                    var meterialthick = thickness.ToString();
                    Debug.WriteLine("meterial:"+meterial);
                    if (meterial.ToLower().Contains("sus"))meterialthick="sus"+ thickness.ToString() ;
                    outputfile = directory + "\\" + "出图" + "\\" + "工程图" + "\\" +meterialthick;
                    
                }
            }
            
            var boolstatus =
                swModel.Extension.SetUserPreferenceDouble((int)swUserPreferenceDoubleValue_e.swDetailingArrowWidth, 0,
                    0.002);
            boolstatus =
                swModel.Extension.SetUserPreferenceDouble((int)swUserPreferenceDoubleValue_e.swDetailingArrowHeight,
                    0, 0.0005);
            boolstatus =
                swModel.Extension.SetUserPreferenceDouble((int)swUserPreferenceDoubleValue_e.swDetailingArrowLength,
                    0, 0.0031);
           
            if (!Directory.Exists(outputfile))
            {
                Directory.CreateDirectory(outputfile);
            }
            string dwgFileName = outputfile + "\\" + Path.GetFileNameWithoutExtension(fullpath) + ".dwg";
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