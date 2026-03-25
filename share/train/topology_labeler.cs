using System;
using System.Collections.Generic;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace tools
{
    /// <summary>
    /// 零件拓扑标注器 - 提供交互式标注功能
    /// </summary>
    public static class TopologyLabeler
    {
        private static TopologyDatabase? _database;
        private static readonly string DefaultDbPath = "topology_labels.db";

        /// <summary>
        /// 初始化数据库
        /// </summary>
        public static void Initialize(string? dbPath = null)
        {
            _database = new TopologyDatabase(dbPath ?? DefaultDbPath);
            Console.WriteLine("✓ 拓扑标注系统已初始化");
        }

        /// <summary>
        /// 标注当前打开的零件
        /// </summary>
        public static void LabelCurrentPart(ModelDoc2 swModel, int wlIterations = 1)
        {
            if (swModel == null)
            {
                Console.WriteLine("× 错误：没有打开的活动文档");
                return;
            }

            if (_database == null)
            {
                Initialize();
            }

            try
            {
                // 构建拓扑图
                Console.WriteLine("\n=== 构建零件拓扑图 ===");
                var graph = FaceGraphBuilder.BuildGraph(swModel);
                
                if (graph == null || graph.Nodes.Count == 0)
                {
                    Console.WriteLine("× 无法构建拓扑图");
                    return;
                }

                // 执行 WL 迭代
                Console.WriteLine($"\n=== 执行 WL 迭代 ({wlIterations} 次) ===");
                var wlFrequencies = WLGraphKernel.PerformWLIterations(graph, wlIterations);

                // 获取零件信息
                string partName = swModel.GetTitle();
                string fullPath = swModel.GetPathName();

                // 存储到数据库
                Console.WriteLine("\n=== 存储到数据库 ===");
                int partId = _database!.UpsertPart(partName, fullPath, wlFrequencies);

                // 显示现有标注
                var existingLabels = _database.GetPartLabels(partId);
                if (existingLabels.Count > 0)
                {
                    Console.WriteLine("\n=== 现有标注 ===");
                    foreach (var labelCategory in existingLabels)
                    {
                        foreach (var (value, confidence, notes) in labelCategory.Value)
                        {
                            Console.WriteLine($"  {labelCategory.Key}: {value} (置信度：{confidence}) - {notes}");
                        }
                    }
                }

                // 提示用户输入新标注
                Console.WriteLine("\n=== 添加标注 ===");
                
                // 只标注一个类别
                Console.Write("标注类别 > ");
                string? category = Console.ReadLine()?.Trim();
                                    
                if (!string.IsNullOrEmpty(category))
                {
                    // 添加标注（值设为空）
                    _database!.AddLabel(partId, category, value: "", confidence: 1.0, notes: "");
                    Console.WriteLine($"✓ 标注已添加：{category}");
                }
                else
                {
                    Console.WriteLine("✓ 标注完成");
                }

                // 显示统计信息
                ShowStatistics();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"× 标注过程出错：{ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        /// <summary>
        /// 查看所有已标注的零件
        /// </summary>
        public static void ViewAllParts()
        {
            if (_database == null)
            {
                Initialize();
            }

            var parts = _database?.GetAllPartsWithLabels();
            
            if (parts == null || parts.Count == 0)
            {
                Console.WriteLine("数据库中暂无零件");
                return;
            }

            Console.WriteLine($"\n=== 已标注零件 ({parts.Count} 个) ===\n");
            
            foreach (var (partId, partName, labels) in parts)
            {
                Console.WriteLine($"[{partId}] {partName}");
                
                if (labels.Count > 0)
                {
                    Console.WriteLine("    标注:");
                    foreach (var label in labels)
                    {
                        Console.WriteLine($"      {label.Key}: {label.Value}");
                    }
                }
                Console.WriteLine();
            }
        }

        /// <summary>
        /// 按标注类别搜索零件
        /// </summary>
        public static void SearchByCategory(string category, string? value = null)
        {
            if (_database == null)
            {
                Initialize();
            }

            var results = _database?.SearchByLabel(category, value);
            
            if (results == null || results.Count == 0)
            {
                Console.WriteLine($"未找到符合条件的零件");
                return;
            }

            Console.WriteLine($"\n=== 搜索结果：{category}" + (value != null ? $" = {value}" : "") + " ===\n");
            
            foreach (var (partId, partName, labelValue) in results)
            {
                Console.WriteLine($"[{partId}] {partName} => {labelValue}");
            }
        }

        /// <summary>
        /// 显示数据库统计信息
        /// </summary>
        public static void ShowStatistics()
        {
            if (_database == null)
            {
                Initialize();
            }

            var stats = _database?.GetStatistics();
            
            if (stats == null) return;
            
            var (partCount, labelCount, categoryStats) = stats.Value;
            
            Console.WriteLine($"\n=== 数据库统计 ===");
            Console.WriteLine($"零件总数：{partCount}");
            Console.WriteLine($"标注总数：{labelCount}");
            Console.WriteLine("标注类别分布:");
            
            foreach (var stat in categoryStats)
            {
                Console.WriteLine($"  {stat.Key}: {stat.Value} 个标注");
            }
            Console.WriteLine();
        }

        /// <summary>
        /// 导出零件的 WL 结果
        /// </summary>
        public static void ExportPartWL(int partId)
        {
            if (_database == null)
            {
                Initialize();
            }

            try
            {
                string json = _database!.ExportWLResult(partId);
                string fileName = $"part_{partId}_wl_export.json";
                System.IO.File.WriteAllText(fileName, json);
                Console.WriteLine($"✓ 已导出 WL 结果到：{fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"× 导出失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 批量标注文件夹中的零件
        /// </summary>
        public static void BatchLabelFolder(string folderPath, SldWorks swApp)
        {
            if (_database == null)
            {
                Initialize();
            }

            Console.WriteLine($"\n=== 批量标注文件夹：{folderPath} ===\n");
            
            var sldprtFiles = System.IO.Directory.GetFiles(folderPath, "*.sldprt");
            var sldasmFiles = System.IO.Directory.GetFiles(folderPath, "*.sldasm");
            var allFiles = sldprtFiles.Concat(sldasmFiles).ToArray();

            if (allFiles.Length == 0)
            {
                Console.WriteLine("× 未找到 SolidWorks 零件文件");
                return;
            }

            Console.WriteLine($"找到 {allFiles.Length} 个文件\n");

            foreach (var filePath in allFiles)
            {
                Console.WriteLine($"\n处理：{System.IO.Path.GetFileName(filePath)}");
                
                // 打开文档
                int errors = 0;
                int warnings = 0;
                ModelDoc2 model = swApp.OpenDoc6(filePath, 
                    (int)swDocumentTypes_e.swDocPART, 
                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent, 
                    "", ref errors, ref warnings);

                if (model == null)
                {
                    Console.WriteLine($"× 无法打开文件：{filePath}");
                    continue;
                }

                try
                {
                    // 自动计算并存储 WL 特征
                    var graph = FaceGraphBuilder.BuildGraph(model);
                    if (graph != null && graph.Nodes.Count > 0)
                    {
                        var wlFrequencies = WLGraphKernel.PerformWLIterations(graph, 1);
                        string partName = model.GetTitle();
                        
                        int partId = _database!.UpsertPart(partName, filePath, wlFrequencies);
                        Console.WriteLine($"✓ 已存储：{partName} ({graph.Nodes.Count} 个面)");
                    }
                    else
                    {
                        Console.WriteLine($"○ 跳过（无实体）：{model.GetTitle()}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"× 处理失败：{ex.Message}");
                }
                finally
                {
                    // 关闭文档
                    swApp.CloseDoc(model.GetTitle());
                }
            }

            Console.WriteLine("\n=== 批量处理完成 ===");
            ShowStatistics();
            
            Console.WriteLine("\n提示：使用 'view_parts' 命令查看已标注的零件，然后手动添加标注");
        }
    }
}
