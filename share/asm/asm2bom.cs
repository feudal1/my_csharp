//https://help.solidworks.com/2023/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ITableAnnotation_members.html
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Diagnostics;
using System.Threading.Tasks;

namespace tools
{
    public class asm2bom
    {
        static public async Task<int> run(SldWorks swApp, ModelDoc2 swModel)
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
                var swBOMAnnotation2 = swBOMAnnotation;
                var colunmname = swTableAnnotation.GetColumnTitle2(2, false);
                var deleteresult=swTableAnnotation.DeleteColumn2(2, false);
                swModel.EditRebuild3();
                Console.WriteLine($"deleteresult:{deleteresult},colunmname{colunmname}");
                swTableAnnotation.InsertColumn2((int)swTableItemInsertPosition_e.swTableItemInsertPosition_After,2,"生产类型",(int)swInsertTableColumnWidthStyle_e.swInsertColumn_DefaultWidth);
                var count = swTableAnnotation.RowCount;
                TopologyLabeler.Initialize();
                var database = TopologyLabeler.GetDatabase();
              
                // 从后向前遍历，避免插入行后影响索引
                int currentRowCount = count;
                for (int i = currentRowCount; i >= 1; i--)
                {
                    var cellText = swTableAnnotation.get_Text(i, 1);
                    if (cellText == null)
                    {
                        Console.WriteLine($"警告：第 {i} 行第 1 列的单元格为空，跳过。");
                        continue;
                    }
                    var partname = cellText.Trim();
                    Debug.WriteLine($"partname:{partname}");
                    
                    // 获取所有 body 及其标签，格式：bodyName1,label1;bodyName2,label2
                    string labelsString = database!.GetLabelsByPartName(partname);
                                        
                    if (string.IsNullOrEmpty(labelsString))
                    {
                            swTableAnnotation.set_Text(i, 3, "");
                       
                        
                        continue;
                    }
                                        
                    // 分割每个 body 的标注（用分号分隔）
                    string[] bodyLabelPairs = labelsString.Split(';');
                    
                    if (bodyLabelPairs.Length == 1)
                    {
                        // 设置当前行的信息（第一个 body）
                        string firstPair = bodyLabelPairs[0];
                        string[] parts = firstPair.Split(',');
                        string bodyName = parts[0];
                       string label = parts[1];
                        swTableAnnotation.set_Text(i, 3, label);
                    
                    }
                  
            
                }

               
                
                string excelpath = swModel.GetPathName().Replace("SLDASM", "xlsx");
                
                swBOMAnnotation.SaveAsExcel(excelpath, false, true);
            swBOMAnnotation=swBOMAnnotation2;
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