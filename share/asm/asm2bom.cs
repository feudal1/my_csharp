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
        static public async Task<int> run(SldWorks swApp, ModelDoc2 swModel,bool issheetmeet)
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
                swTableAnnotation.InsertColumn2((int)swTableItemInsertPosition_e.swTableItemInsertPosition_After,2,"规格尺寸",(int)swInsertTableColumnWidthStyle_e.swInsertColumn_DefaultWidth);
                swTableAnnotation.SetColumnTitle(2,"单套数量");
                swTableAnnotation.InsertColumn2((int)swTableItemInsertPosition_e.swTableItemInsertPosition_After,3,"是否出图",(int)swInsertTableColumnWidthStyle_e.swInsertColumn_DefaultWidth);
                swTableAnnotation.InsertColumn2((int)swTableItemInsertPosition_e.swTableItemInsertPosition_After,4,"总数",(int)swInsertTableColumnWidthStyle_e.swInsertColumn_DefaultWidth);
                var count = swTableAnnotation.RowCount;
   
                // 获取零件文档 - 使用 GetComponents 方法遍历所有组件来查找
                AssemblyDoc swAssembly = (AssemblyDoc)swModel;
                Component2 targetComponent = null;
                                
                // 获取所有组件
                object[] allComponents = (object[])swAssembly.GetComponents(false);
                
                // 创建字典用于累加每个零件的数量
                Dictionary<string, int> partCountDict = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                // 正向遍历BOM表行
                int currentRowCount = count;
                int asmfactor = 0;
                for (int i = 1; i <= currentRowCount; i++)
                {
                
                    var cellText = swTableAnnotation.get_Text(i, 1);
                    if (cellText == null)
                    {
                        Console.WriteLine($"警告：第 {i} 行第 1 列的单元格为空，跳过。");
                        continue;
                    }
                    var partname = cellText.Trim();
                    swTableAnnotation.set_Text(i, 1, partname.Replace("=", ""));
                    Debug.WriteLine($"partname:{partname}");


                    
                    // 检查是否存在对应的 DWG 文件
                    bool dwgExists = false;
                    if (!string.IsNullOrEmpty(directory) && Directory.Exists(directory))
                    {
                        dwgExists = FindDwgFile(directory, partname);
                    }
                    
         
                        if (dwgExists)
                        {
                            swTableAnnotation.set_Text(i, 4, "已出图");
                        }
                        // 安全地获取并解析数量值
                        string itemCountText = swTableAnnotation.get_Text(i, 2);
                        int itemcount = 0;
                        if (!string.IsNullOrEmpty(itemCountText))
                        {
                            if (!int.TryParse(itemCountText, out itemcount))
                            {
                                Console.WriteLine($"警告：{partname}第 {i} 行第 2 列的数量值 '{itemCountText}' 无法转换为整数，使用默认值 0。");
                                itemcount = 0;
                            }
                        }
                        else
                        {
                            Console.WriteLine($"警告：第 {i} 行第 2 列的数量值为空，使用默认值 0。");
                        }
                 
    
     
                    
                            try
                            {
                        
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
                                       
                                        if (targetComponent == null)
                                {
                                    continue; // 没找到匹配的组件，跳过当前BOM行
                                }
                                ModelDoc2 partDoc = (ModelDoc2)targetComponent.GetModelDoc2();
                                


                                if (partDoc != null && partDoc.GetType() == (int)swDocumentTypes_e.swDocPART)
                                {
                                    if (!issheetmeet)
                                    {
                                        PartDoc part = (PartDoc)partDoc;

                                        var dimensions = PartDimensionHelper.GetPartDimensions(part);

                                        // 将尺寸格式化为字符串并添加到BOM表
                                        string dimensionStr =
                                            $"{dimensions.length}x{dimensions.width}x{dimensions.height}";
                                        if (partname != "脚杯座")
                                            dimensionStr = dimensionStr.Replace("40x60", "2.0x40x60")
                                                .Replace("60x40", "2.0x40x60").Replace("40x40", "2.0x40x40");
                                        if (partname.Contains("方管"))
                                            Console.WriteLine($"管件 '{partname}' 尺寸: {dimensionStr}");
                                        swTableAnnotation.set_Text(i, 3, dimensionStr);

                                        Debug.WriteLine($"管件 '{partname}' 尺寸: {dimensionStr}");
                                        break; // 找到第一个匹配的就跳出内层循环
                                    }
                                    else
                                    { 
                                        PartDoc part = (PartDoc)partDoc;
                                            
                                            // 检查是否为钣金件
                                            Feature swFeature = (Feature)part.FirstFeature();
                                            bool isSheetMetal = false;
                                            double thickness = 0;
                                            double boundingBoxLength = 0;
                                            double boundingBoxWidth = 0;
                                            
                                            while (swFeature != null)
                                            {
                                               
                                                
                                                // 查找SolidBodyFolder以获取切割清单信息
                                                if (swFeature.GetTypeName2() == "SolidBodyFolder")
                                                {
                                                    BodyFolder swBodyFolder = (BodyFolder)swFeature.GetSpecificFeature2();
                                                    swBodyFolder.SetAutomaticCutList(true);
                                                    swBodyFolder.SetAutomaticUpdate(true);
                                                    
                                                    Feature subfeat = (Feature)swFeature.GetFirstSubFeature();
                                                    
                                                    while (subfeat != null)
                                                    {
                                                        IBodyFolder solidBodyFolder = (IBodyFolder)subfeat.GetSpecificFeature2();
                                                        
                                                        var manger = subfeat.CustomPropertyManager;
                                                        object vPropNames = null;
                                                        object vPropTypes = null;
                                                        object vPropValues = null;
                                                        object resolved = null;
                                                        object linkProp = null;
                                                        
                                                        manger.GetAll3(ref vPropNames, ref vPropTypes, ref vPropValues, ref resolved, ref linkProp);
                                                        
                                                        if (vPropValues != null && vPropNames != null)
                                                        {
                                                            string[] propValues = (string[])vPropValues;
                                                            string[] propNames = (string[])vPropNames;
                                                            
                                                            for (int j = 0; j < propNames.Length; j++)
                                                            {
                                                                if ((propNames[j] == "边界框长度" || propNames[j] == "Bounding Box Length") && 
                                                                    double.TryParse(propValues[j], out double length))
                                                                {
                                                                    boundingBoxLength = length;
                                                                }
                                                                if ((propNames[j] == "边界框宽度" || propNames[j] == "Bounding Box Width") && 
                                                                    double.TryParse(propValues[j], out double width))
                                                                {
                                                                    boundingBoxWidth = width;
                                                                }
                                                                if ((propNames[j] == "钣金厚度" || propNames[j] == "Thickness") && 
                                                                    double.TryParse(propValues[j], out double thicknessValue))
                                                                {
                                                                    thickness = thicknessValue;
                                                                    if (thicknessValue > 0.1) 
                                                                    {
                                                                        isSheetMetal = true;
                                                                        
                                                                        string material="SPPC " + thickness.ToString("F2");
                                                                        string  materialdataname = "";
                                                                        string materialname=part.GetMaterialPropertyName2("Default",
                                                                            out materialdataname);
                                                                        if(materialname.Contains("不锈钢")&&materialname.Contains("sus"))material="SUS " + thickness.ToString("F2");
                                                                      partDoc.DeleteCustomInfo2("", "材料");
                                                                        bool materialResult=partDoc.AddCustomInfo2("材料", (int)swCustomInfoType_e.swCustomInfoText, material);
                                                                        Console.WriteLine($"添加材料自定义信息结果: {materialResult},partname:{partname},material:{material}");
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        
                                                        subfeat = (Feature)subfeat.GetNextSubFeature();
                                                    }
                                                }
                                                
                                                swFeature = (Feature)swFeature.GetNextFeature();
                                            }
                                            
                                            string dimensionStr = "";
                                            
                                            if (isSheetMetal && boundingBoxLength > 0 && boundingBoxWidth > 0)
                                            {
                                                // 使用切割清单中的边界框尺寸作为下料尺寸
                                                double maxLength = Math.Max(boundingBoxLength, boundingBoxWidth);
                                                double minLength = Math.Min(boundingBoxLength, boundingBoxWidth);
                                                dimensionStr = $"{maxLength}x{minLength}x{thickness}";
                                                
                                                Debug.WriteLine($"钣金件 '{partname}' 下料尺寸: {dimensionStr}");
                                            }
                                            else if (isSheetMetal)
                                            {
                                                // 如果没有找到切割清单信息，使用普通边界框
                                                var dimensions = PartDimensionHelper.GetPartDimensions(part);
                                                dimensionStr = $"{dimensions.length:F1}x{dimensions.width:F1}x{thickness:F1}";
                                            }
                                            else
                                            {
                                                // 非钣金件,使用普通尺寸
                                                var dimensions = PartDimensionHelper.GetPartDimensions(part);
                                                dimensionStr = $"{dimensions.length:F0}x{dimensions.width:F0}x{dimensions.height:F0}";
                                                       partDoc.DeleteCustomInfo2("", "规格");
                                                                        bool materialResult=partDoc.AddCustomInfo2("规格", (int)swCustomInfoType_e.swCustomInfoText, dimensionStr);
                                            }
                                            
                                            swTableAnnotation.set_Text(i, 3, dimensionStr);
                                    }
                                
                                    itemcount = itemcount*asmfactor;
                                    } 
                             
                                    else  
                                    {

                                        asmfactor = itemcount;
                                    }
                                    partDoc.DeleteCustomInfo2("", "数量");
                                    // 使用字典中累加后的数量值
                                    int accumulatedCount = partCountDict.ContainsKey(partname) ? partCountDict[partname] : itemcount;
                                    bool result=partDoc.AddCustomInfo2("数量", (int)swCustomInfoType_e.swCustomInfoText, accumulatedCount.ToString());
                                    Console.WriteLine($"添加数量自定义信息结果: {result},partcount:{accumulatedCount},partname:{partname}");

                                    break;
                                    }
                                }
                                
                                
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"获取管件 '{partname}' 尺寸时出错: {ex.Message}");
                            }
                    
                    
                       
                    
                            // 累加零件数量到字典
                            if (!partCountDict.ContainsKey(partname))
                            {
                                partCountDict[partname] = 0;
                            }
                            partCountDict[partname] += itemcount;
                    
                            swTableAnnotation.set_Text( i, 2, itemcount.ToString());
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