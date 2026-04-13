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
                string fullPath = swModel.GetPathName();

                if (string.IsNullOrEmpty(fullPath))
                {
                    Console.WriteLine("错误：文档尚未保存，请先保存文件。");
                    return -1;
                }
                string? directory = Path.GetDirectoryName(fullPath) + "\\" + "出图";

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
                   (int)swNumberingType_e.swNumberingType_None, 
                    true
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
                swTableAnnotation.InsertColumn2((int)swTableItemInsertPosition_e.swTableItemInsertPosition_After,3,"是否出图",(int)swInsertTableColumnWidthStyle_e.swInsertColumn_DefaultWidth);
                swTableAnnotation.InsertColumn2((int)swTableItemInsertPosition_e.swTableItemInsertPosition_After,4,"尺寸",(int)swInsertTableColumnWidthStyle_e.swInsertColumn_DefaultWidth);
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
                    var partnumber=swTableAnnotation.get_Text(i, 0);
                 
                    // 检查是否存在对应的 DWG 文件
                    bool dwgExists = false;
                    if (!string.IsNullOrEmpty(directory) && Directory.Exists(directory))
                    {
                        dwgExists = FindDwgFile(directory, partname);
                    }
                    
                    // 获取所有 body 及其标签，格式：bodyName1,label1;bodyName2,label2
                    string labelsString = database!.GetLabelsByPartName(partname);
                                        
                 
                        if (dwgExists)
                        {
                            swTableAnnotation.set_Text(i, 4, "已出图");
                        }
                      
                 
                                         // 检查 labelsString 是否有效
                    if (string.IsNullOrEmpty(labelsString))
                    {
                        Debug.WriteLine($"警告：零件 '{partname}' 没有找到拓扑标注信息。");
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
                       string label =  parts[1];
                        swTableAnnotation.set_Text(i, 3,  label);
                        
                        // 如果标签是"管件"，获取零件尺寸并添加到BOM表
                        if (label == "管件")
                        {
                            try
                            {
                                // 获取零件文档 - 使用 GetComponents 方法遍历所有组件来查找
                                AssemblyDoc swAssembly = (AssemblyDoc)swModel;
                                Component2 targetComponent = null;
                                
                                // 获取所有组件
                                object[] allComponents = (object[])swAssembly.GetComponents(false);
                                
                                // 遍历所有组件，按名称查找目标组件
                                foreach (object compObj in allComponents)
                                {
                                    Component2 component = (Component2)compObj;
                                    string componentName = component.Name2;
                                    
                                    // 去掉"/"号及之前的文字，只保留后面的部分
                                    int slashIndex = componentName.LastIndexOf('/');
                                    if (slashIndex >= 0 && slashIndex < componentName.Length - 1)
                                    {
                                        componentName = componentName.Substring(slashIndex + 1);
                                    }
                                    
                                    // 去掉末尾的"-数字"部分（例如："零件名-1" -> "零件名"）
                                    int lastDashIndex = componentName.LastIndexOf('-');
                                    if (lastDashIndex > 0 && lastDashIndex < componentName.Length - 1)
                                    {
                                        string suffix = componentName.Substring(lastDashIndex + 1);
                                        // 检查后缀是否为纯数字
                                        if (int.TryParse(suffix, out _))
                                        {
                                            componentName = componentName.Substring(0, lastDashIndex);
                                        }
                                    }
                                    
                                    // 比较组件名称（不区分大小写）
                                    if (componentName.Equals(partname, StringComparison.OrdinalIgnoreCase))
                                    {
                                        targetComponent = component;
                                        break; // 找到第一个匹配的就跳出内层循环
                                    }
                                }
                                
                                if (targetComponent == null)
                                {
                                    continue; // 没找到匹配的组件，跳过当前BOM行
                                }
                                
                                ModelDoc2 partDoc = (ModelDoc2)targetComponent.GetModelDoc2();
                                if (partDoc != null && partDoc.GetType() == (int)swDocumentTypes_e.swDocPART)
                                {
                                    PartDoc part = (PartDoc)partDoc;
                                    var dimensions = PartDimensionHelper.GetPartDimensions(part);
                                    
                                    // 将尺寸格式化为字符串并添加到BOM表
                                    string dimensionStr = $"{dimensions.length}x{dimensions.width}x{dimensions.height}";
                                    swTableAnnotation.set_Text(i, 5, dimensionStr);
                                    
                                    Debug.WriteLine($"管件 '{partname}' 尺寸: {dimensionStr}");
                                }
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"获取管件 '{partname}' 尺寸时出错: {ex.Message}");
                            }
                        }
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

        /// <summary>
        /// 在文件夹及其子文件夹中查找是否存在与零件名匹配的 DWG 文件
        /// </summary>
        /// <param name="directory">要搜索的目录</param>
        /// <param name="partName">零件名称</param>
        /// <returns>如果找到返回 true，否则返回 false</returns>
        static private bool FindDwgFile(string directory, string partName)
        {
            try
            {
                // 搜索当前目录和所有子目录中的 DWG 文件（包括 sus、CNC 等材质文件夹）
                string[] dwgFiles = Directory.GetFiles(directory, "*.dwg", SearchOption.AllDirectories);
                
                foreach (string dwgFile in dwgFiles)
                {
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(dwgFile);
                    
                    // 精确匹配（不区分大小写）
                    if (fileNameWithoutExt.Equals(partName, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                    
                    // 宽松匹配：去除空格、连字符等特殊字符后比较
                    string cleanPartName = partName.Replace(" ", "").Replace("-", "").Replace("_", "");
                    string cleanFileName = fileNameWithoutExt.Replace(" ", "").Replace("-", "").Replace("_", "");
                    if (cleanFileName.Equals(cleanPartName, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
                
                return false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"搜索 DWG 文件时出错：{ex.Message}");
                return false;
            }
        }
    }
}