using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace tools
{
    public class asm2do
    {
        static private Dictionary<string, int> docnameCount = new Dictionary<string, int>();
        static private Dictionary<string, int> dwgExportCount = new Dictionary<string, int>();
        static private readonly object _swApiLock = new object();
        static private string[] currentFilterKeywords = null;

        public delegate int SolidWorksAction( ModelDoc2 swModel,SldWorks swApp=null);

        static public string[]? run(SldWorks swApp, ModelDoc2 swModel, SolidWorksAction action, string[] filterKeywords = null)
        {
            try
            {
                // 清空上次的统计结果
                docnameCount.Clear();
                dwgExportCount.Clear();

                if (swModel == null)
                {
                    Console.WriteLine("错误：没有打开任何文档。");
                    return null;
                }

                // 确保是装配体
                if (swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    Console.WriteLine("错误：当前文档不是装配体。");
                    return null;
                }

                AssemblyDoc swAssembly = (AssemblyDoc)swModel;

                // 获取顶层组件
                object[] topComponents = (object[])swAssembly.GetComponents(true);

                Console.WriteLine($"正在扫描装配体... 共 {topComponents.Length} 个顶层组件");

                // 如果指定了筛选关键字，输出提示信息
                if (filterKeywords != null && filterKeywords.Length > 0)
                {
                    Console.WriteLine($"筛选关键字: {string.Join(", ", filterKeywords)}");
                    currentFilterKeywords = filterKeywords;
                }
                else
                {
                    currentFilterKeywords = null;
                }

                // 先统计需要处理的零件总数
                int totalPartsToProcess = CountUniqueParts(swAssembly, filterKeywords);
                int processedCount = 0;

                Console.WriteLine($"\n开始处理，共需处理 {totalPartsToProcess} 个唯一零件...\n");

                // 单线程顺序处理
                foreach (object compObj in topComponents)
                {
                   
                    Component2 topComp = (Component2)compObj;
                    Console.WriteLine($"正在处理 {topComp.Name2}...");
                    
                    // 获取顶层组件的模型文档并应用相同的清理逻辑
                    ModelDoc2 doc = (ModelDoc2)topComp.GetModelDoc2();
                    if (doc != null)
                    {
                        doc.Visible = true;
                        string? pathName = doc.GetPathName();
                        if (!string.IsNullOrEmpty(pathName))
                        {
                            string docname = pathName.Replace(@"\\", "/");
                            string fileName = Path.GetFileName(docname);
                            string directory = Path.GetDirectoryName(docname);
                            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                            string extension = Path.GetExtension(fileName);
                            
                            // 去掉文件名末尾的"-数字"部分（例如："零件名-1" -> "零件名"）
                            int lastDashIndex = fileNameWithoutExt.LastIndexOf('-');
                            if (lastDashIndex > 0 && lastDashIndex < fileNameWithoutExt.Length - 1)
                            {
                                string suffix = fileNameWithoutExt.Substring(lastDashIndex + 1);
                                // 检查后缀是否为纯数字
                                if (int.TryParse(suffix, out _))
                                {
                                    fileNameWithoutExt = fileNameWithoutExt.Substring(0, lastDashIndex);
                                }
                            }
                            
                            // 重新组合完整的文件路径
                            string cleanedFileName = fileNameWithoutExt + extension;
                            string cleanedDocname = Path.Combine(directory, cleanedFileName).Replace(@"\\", "/");

                            // 如果有关键字筛选，检查是否匹配
                            if (currentFilterKeywords != null && currentFilterKeywords.Length > 0)
                            {
                                bool shouldProcess = false;
                                foreach (string keyword in currentFilterKeywords)
                                {
                                    if (!string.IsNullOrEmpty(keyword) && fileNameWithoutExt.Contains(keyword))
                                    {
                                        shouldProcess = true;
                                        break;
                                    }
                                }
                                
                                if (!shouldProcess)
                                {
                                    // 不匹配筛选条件，跳过此零件
                                    swApp.CloseDoc(pathName);
                                    continue;
                                }
                            }

                            // 使用清理后的名称添加到字典
                            if (!docnameCount.ContainsKey(cleanedDocname))
                            {
                                docnameCount[cleanedDocname] = 0;
                            }

                            docnameCount[cleanedDocname]++;

                            // 初始化 DWG 计数
                            if (!dwgExportCount.ContainsKey(cleanedDocname))
                            {
                                dwgExportCount[cleanedDocname] = 0;
                            }

                            // 只处理第一次遇到的零件
                            if (docnameCount[cleanedDocname] == 1)
                            {
                                processedCount++;
                                Debug.WriteLine($"[{processedCount}/{totalPartsToProcess}] 正在处理 {cleanedDocname}...");
                                int exportedCount = action.Invoke(doc, swApp);
                                dwgExportCount[cleanedDocname] = exportedCount;
                                Console.WriteLine($"进度: {processedCount}/{totalPartsToProcess} ({(processedCount * 100.0 / totalPartsToProcess):F1}%)");
                            }

                            // 成功处理完后，关闭文档
                            swApp.CloseDoc(pathName);
                        }
                    }
                    
                    TraverseComponent(swApp, topComp, action, totalPartsToProcess, ref processedCount);
                }



                Console.WriteLine("\n========== 零件统计信息 ==========");
                foreach (var kvp in docnameCount.OrderBy(x => x.Key))
                {
                    string partName = Path.GetFileName(kvp.Key);
                    int refCount = kvp.Value;
                    int dwgCount = dwgExportCount.ContainsKey(kvp.Key) ? dwgExportCount[kvp.Key] : 0;
                    Console.WriteLine($"{partName}: 引用 {refCount} 次，导出 {dwgCount} 个 DWG");
                }
                Console.WriteLine($"\n总计：{docnameCount.Count} 个零件，共导出 {dwgExportCount.Values.Sum()} 个 DWG 文件");
                Console.WriteLine("====================================\n");

                List<string> result = new List<string>(docnameCount.Keys);
                return result.ToArray();

            }
            catch (Exception ex)
            {
                Console.WriteLine($"发生错误: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }

            return null;
        }

// 递归遍历组件
        static private void TraverseComponent(SldWorks swApp, Component2 component, SolidWorksAction action, int totalParts, ref int processedCount)
        {
            // 获取子组件
            object[] children = (object[])component.GetChildren();

            if (children != null && children.Length > 0)
            {
                foreach (object childObj in children)
                {
                    Component2 childComp = (Component2)childObj;
                    TraverseComponent(swApp, childComp, action, totalParts, ref processedCount);
                }
            }
            else
            {
                ModelDoc2 doc = (ModelDoc2)component.GetModelDoc2();
                if (doc != null)
                {
                    doc.Visible = true;
                    string? pathName = doc.GetPathName();
                    if (string.IsNullOrEmpty(pathName))
                    {
                        Console.WriteLine($"警告：组件 {component.Name2} 未保存或无法获取路径。");
                        // 即使没有路径，也要尝试关闭文档以释放资源
                        swApp.CloseDoc("");
                        return;
                        
                    }

                    string docname = pathName.Replace(@"\\", "/");

              
                    string fileName = Path.GetFileName(docname);
                    string directory = Path.GetDirectoryName(docname);
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                    string extension = Path.GetExtension(fileName);
                    
                    // 去掉文件名末尾的"-数字"部分（例如："零件名-1" -> "零件名"）
                    int lastDashIndex = fileNameWithoutExt.LastIndexOf('-');
                    if (lastDashIndex > 0 && lastDashIndex < fileNameWithoutExt.Length - 1)
                    {
                        string suffix = fileNameWithoutExt.Substring(lastDashIndex + 1);
                        // 检查后缀是否为纯数字
                        if (int.TryParse(suffix, out _))
                        {
                            fileNameWithoutExt = fileNameWithoutExt.Substring(0, lastDashIndex);
                        }
                    }
          
                    
                    // 重新组合完整的文件路径
                    string cleanedFileName = fileNameWithoutExt + extension;
                    string cleanedDocname = Path.Combine(directory, cleanedFileName).Replace(@"\\", "/");

                    // 如果有关键字筛选，检查是否匹配
                    if (currentFilterKeywords != null && currentFilterKeywords.Length > 0)
                    {
                        bool shouldProcess = false;
                        foreach (string keyword in currentFilterKeywords)
                        {
                            if (!string.IsNullOrEmpty(keyword) && fileNameWithoutExt.Contains(keyword))
                            {
                                shouldProcess = true;
                                break;
                            }
                        }
                        
                        if (!shouldProcess)
                        {
                            // 不匹配筛选条件，跳过此零件
                            swApp.CloseDoc(pathName);
                            return;
                        }
                    }

                    // 使用清理后的名称添加到字典
                    if (!docnameCount.ContainsKey(cleanedDocname))
                    {
                        docnameCount[cleanedDocname] = 0;
                    }

                    docnameCount[cleanedDocname]++;

                    // 初始化 DWG 计数
                    if (!dwgExportCount.ContainsKey(cleanedDocname))
                    {
                        dwgExportCount[cleanedDocname] = 0;
                    }

                    // 只处理第一次遇到的零件
                    if (docnameCount[cleanedDocname] == 1)
                    {
                        processedCount++;
                        Debug.WriteLine($"[{processedCount}/{totalParts}] 正在处理 {cleanedDocname}...");
                        int exportedCount = action.Invoke( doc, swApp);
                        dwgExportCount[cleanedDocname] = exportedCount;
                        Console.WriteLine($"进度: {processedCount}/{totalParts} ({(processedCount * 100.0 / totalParts):F1}%)");
                    }

                    // 成功处理完后，关闭文档
                    swApp.CloseDoc(pathName);



                }
            }
        }

        // 统计需要处理的唯一零件数量
        static private int CountUniqueParts(AssemblyDoc swAssembly, string[] filterKeywords)
        {
            Dictionary<string, bool> uniqueParts = new Dictionary<string, bool>();
            object[] topComponents = (object[])swAssembly.GetComponents(true);
            
            foreach (object compObj in topComponents)
            {
                Component2 topComp = (Component2)compObj;
                CountUniquePartsRecursive(topComp, uniqueParts, filterKeywords);
            }
            
            return uniqueParts.Count;
        }

        static private void CountUniquePartsRecursive(Component2 component, Dictionary<string, bool> uniqueParts, string[] filterKeywords)
        {
            object[] children = (object[])component.GetChildren();

            if (children != null && children.Length > 0)
            {
                foreach (object childObj in children)
                {
                    Component2 childComp = (Component2)childObj;
                    CountUniquePartsRecursive(childComp, uniqueParts, filterKeywords);
                }
            }
            else
            {
                ModelDoc2 doc = (ModelDoc2)component.GetModelDoc2();
                if (doc != null)
                {
                    string? pathName = doc.GetPathName();
                    if (!string.IsNullOrEmpty(pathName))
                    {
                        string fileName = Path.GetFileName(pathName);
                        string fileNameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                        
                        // 去掉文件名末尾的"-数字"部分
                        int lastDashIndex = fileNameWithoutExt.LastIndexOf('-');
                        if (lastDashIndex > 0 && lastDashIndex < fileNameWithoutExt.Length - 1)
                        {
                            string suffix = fileNameWithoutExt.Substring(lastDashIndex + 1);
                            if (int.TryParse(suffix, out _))
                            {
                                fileNameWithoutExt = fileNameWithoutExt.Substring(0, lastDashIndex);
                            }
                        }
                        
                        // 如果有关键字筛选，检查是否匹配
                        if (filterKeywords != null && filterKeywords.Length > 0)
                        {
                            bool shouldInclude = false;
                            foreach (string keyword in filterKeywords)
                            {
                                if (!string.IsNullOrEmpty(keyword) && fileNameWithoutExt.Contains(keyword))
                                {
                                    shouldInclude = true;
                                    break;
                                }
                            }
                            
                            if (!shouldInclude)
                            {
                                return;
                            }
                        }
                        
                        string cleanedDocname = pathName.Replace(@"\\", "/");
                        if (!uniqueParts.ContainsKey(cleanedDocname))
                        {
                            uniqueParts[cleanedDocname] = true;
                        }
                    }
                }
            }
        }
    }
}

