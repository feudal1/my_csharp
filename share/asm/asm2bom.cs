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

                var partname = Path.GetFileNameWithoutExtension(swModel.GetPathName());
                Debug.WriteLine($"正在处理装配体：{partname}");

                ModelDocExtension swModelDocExt = (ModelDocExtension)swModel.Extension;
                BomTableAnnotation swBOMAnnotation;
                TableAnnotation swTableAnnotation;
                string tableTemplate;
                int bomType;

                // 使用默认的 BOM 模板路径
                tableTemplate = "C:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\lang\\chinese-simplified\\bom-standard.sldbomtbt";
                
                // 如果默认模板不存在，尝试英文模板
                if (!System.IO.File.Exists(tableTemplate))
                {
                    tableTemplate = "C:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\lang\\english\\bom-standard.sldbomtbt";
                }

                // 插入缩进式 BOM 表
                bomType = (int)swBomType_e.swBomType_Indented;
                
                swBOMAnnotation = (BomTableAnnotation)swModelDocExt.InsertBomTable3(
                    tableTemplate, 
                    0, 
                    1, 
                    bomType, 
                    "Default", 
                    false, 
                    (int)swNumberingType_e.swNumberingType_Detailed, 
                    true
                );

                if (swBOMAnnotation == null)
                {
                    Console.WriteLine("错误：无法插入 BOM 表。");
                    return -1;
                }

                // 获取 BOM 表注释信息
                swTableAnnotation = (TableAnnotation)swBOMAnnotation;
                
                Console.WriteLine($"\n========== BOM 表信息 ==========");
                Console.WriteLine($"装配体：{partname}");
                
                // 输出列标题信息（通过尝试获取列标题来确认列数）
                int columnCount = 0;
                for (int i = 0; i < 50; i++) // 尝试最多 50 列
                {
                    string columnTitle = swTableAnnotation.GetColumnTitle(i);
                    if (string.IsNullOrEmpty(columnTitle))
                    {
                        break;
                    }
                    columnCount++;
                    Console.WriteLine($"  列 {i}: {columnTitle}");
                }
                
                Console.WriteLine($"BOM 表列数：{columnCount}");
                Console.WriteLine("====================================\n");

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