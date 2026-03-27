using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Diagnostics;

namespace tools
{
    public class asm2bom
    {
        static public int run(SldWorks swApp, ModelDoc2 swModel)
        {
            try
            {
                // 确保是装配体
                if (swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    Console.WriteLine("错误：当前文档不是装配体。");
                    return -1;
                }

            

                ModelDocExtension swModelDocExt = (ModelDocExtension)swModel.Extension;
                BomTableAnnotation swBOMAnnotation;
                TableAnnotation swTableAnnotation;
                string tableTemplate;
                int bomType;

                // 使用默认的 BOM 模板路径
                tableTemplate = "C:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\lang\\chinese-simplified\\bom-standard.sldbomtbt";
                
                string Configuration = swApp.GetActiveConfigurationName(swModel.GetPathName());
                // 插入缩进式 BOM 表
                bomType = (int)swBomType_e.swBomType_Indented;
                
                swBOMAnnotation = (BomTableAnnotation)swModelDocExt.InsertBomTable3(
                    tableTemplate, 
                    1260, 
                    556, 
                    bomType, 
                    Configuration, 
                    true, 
                    0, 
                    false
                );

                if (swBOMAnnotation == null)
                {
                    Console.WriteLine("错误：无法插入 BOM 表。");
                    return -1;
                }

                // 获取 BOM 表注释信息
                swTableAnnotation = (TableAnnotation)swBOMAnnotation;



                string excelpath = swModel.GetPathName().Replace("SLDASM", "xlsx");
                swBOMAnnotation.SaveAsExcel(excelpath, true, true);
                
                // 启动 Excel 文件
                ProcessStartInfo startInfo = new ProcessStartInfo(excelpath)
                {
                    UseShellExecute = true
                };
                Process.Start(startInfo);
                
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"发生错误：{ex.Message}");
                Console.WriteLine(ex.StackTrace);
                return -1;
            }
        }
    }
}