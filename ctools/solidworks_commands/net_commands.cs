using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using tools;
using cad_tools;
using Newtonsoft.Json;

namespace tools
{
    partial class Program
    {
        // ========== 拓扑标注相关命令 ==========
        
        [Command("label", Description = "标注当前零件的拓扑图。用法：label [body 名称] [值] - 精确匹配 body 名称并标注。示例：label 凸台拉伸 1 管件、label 法兰盘 钣金件", Parameters = "[body 名称] [值]", Group = "train")]

        static void LabelPart(string[] args)
        {
            if (Program.SwModel == null)
            {
                Console.WriteLine("错误：请先打开一个零件文档");
                return;
            }

            Console.WriteLine("=== 零件拓扑标注 ===\n");
            
            // 构建所有 Body 的拓扑图并执行 WL 迭代
            var graphs = FaceGraphBuilder.BuildGraphs(Program.SwModel);
            
            if (graphs.Count == 0)
            {
                Console.WriteLine("× 无法构建拓扑图");
                return;
            }

            // 获取零件信息
            string partName = Program.SwModel.GetTitle();
            string fullPath = Program.SwModel.GetPathName();

            // 初始化数据库并存储到数据库
            TopologyLabeler.Initialize();
            var database = TopologyLabeler.GetDatabase();
            var bodyIds = database!.UpsertPartWithBodies(partName, fullPath, graphs);

            // 显示所有 body 信息（带索引）
            Console.WriteLine($"\n零件包含 {bodyIds.Count} 个 body:");
            for (int i = 0; i < bodyIds.Count; i++)
            {
                Console.WriteLine($"  [{i}] {graphs[i].BodyName} ({graphs[i].Nodes.Count} 个面)");
            }

            // 检查参数
            if (args.Length < 2)
            {
                Console.WriteLine("\n=== 用法 ===");
                Console.WriteLine("label [body 名称] [值]");
                Console.WriteLine("示例:");
                Console.WriteLine("  label 凸台拉伸 1 管件");
                Console.WriteLine("  label 法兰盘 钣金件");
                Console.WriteLine("  label 底座 框架");
                return;
            }

            string bodyName = args[0];
            string value = args[1];

            // 精确匹配（忽略大小写和首尾空格）
            int targetIndex = -1;
            string normalizedInput = bodyName.Trim().ToLower();
            for (int i = 0; i < graphs.Count; i++)
            {
                if (graphs[i].BodyName.Trim().ToLower() == normalizedInput)
                {
                    targetIndex = i;
                    break;
                }
            }

            if (targetIndex == -1)
            {
                Console.WriteLine($"\n× 错误：未找到名为 '{bodyName}' 的 body");
                Console.WriteLine("\n可用的 body 名称有:");
                for (int i = 0; i < graphs.Count; i++)
                {
                    Console.WriteLine($"  [{i}] {graphs[i].BodyName}");
                }
                Console.WriteLine("\n提示：请按正确的 body 名称重新标注，可以从上面列表中复制。");
                return;
            }

            // 标注指定的 body
            int targetBodyId = bodyIds[targetIndex];
            AddAndShowLabel(database, targetBodyId, graphs[targetIndex], value, targetIndex);

            // 显示统计信息
            TopologyLabeler.ShowStatistics();
        }

        /// <summary>
        /// 添加标注并显示结果
        /// </summary>
        static void AddAndShowLabel(TopologyDatabase database, int bodyId, BodyGraph graph, string value, int bodyIndex)
        {
            // 使用 body 名称作为类别（自动推断）
            string category = graph.BodyName;
                    
            database.AddLabel(bodyId, category, value, confidence: 1.0, notes: "LLM 自动标注");
                    
            Console.WriteLine($"\n✓ 标注已添加:");
            Console.WriteLine($"  Body: {graph.BodyName}");
            Console.WriteLine($"  值：{value}");
        
            // 显示该 body 的所有标注
            var allLabels = database.GetBodyLabels(bodyId);
            if (allLabels.Count > 0)
            {
                Console.WriteLine($"\n  Body [{bodyIndex}] 的所有标注:");
                foreach (var labelCategory in allLabels)
                {
                    foreach (var labelData in labelCategory.Value)
                    {
                        Console.WriteLine($"    {labelCategory.Key}: {labelData.Item1} (置信度：{labelData.Item2}) - {labelData.Item3}");
                    }
                }
            }
        }
        
        /// <summary>
        /// 标注所有 body 为同一个值
        /// </summary>
        [Command("label_all", Description = "标注当前零件的所有 body 为同一个标注。用法：label_all [值] - 无需指定 body 名称，将所有 body 标注为同一类别。示例：label_all 管件、label_all 钣金件、标注当前文件为 xx 也用此方法", Parameters = "[值]", Group = "train")]
        static void LabelAllBodies(string[] args)
        {
            if (Program.SwModel == null)
            {
                Console.WriteLine("错误：请先打开一个零件文档");
                return;
            }
        
            Console.WriteLine("=== 批量标注所有 Body ===\n");
                    
            // 构建所有 Body 的拓扑图
            var graphs = FaceGraphBuilder.BuildGraphs(Program.SwModel);
                    
            if (graphs.Count == 0)
            {
                Console.WriteLine("× 无法构建拓扑图");
                return;
            }
        
            // 获取零件信息
            string partName = Program.SwModel.GetTitle();
            string fullPath = Program.SwModel.GetPathName();
        
            // 初始化数据库并存储到数据库
            TopologyLabeler.Initialize();
            var database = TopologyLabeler.GetDatabase();
            var bodyIds = database!.UpsertPartWithBodies(partName, fullPath, graphs);
        
            // 检查参数
            if (args.Length < 1)
            {
                Console.WriteLine("\n=== 用法 ===");
                Console.WriteLine("label_all [值]");
                Console.WriteLine("示例:");
                Console.WriteLine("  label_all 管件");
                Console.WriteLine("  label_all 钣金件");
                Console.WriteLine("  label_all 结构件");
                return;
            }
        
            string value = args[0];
        
            // 标注所有 body
            Console.WriteLine($"\n正在标注 {bodyIds.Count} 个 body 为 '{value}'...\n");
                    
            for (int i = 0; i < bodyIds.Count; i++)
            {
                string category = $"{partName}_Body{i}"; // 使用零件名 + 索引作为类别
                database.AddLabel(bodyIds[i], category, value, confidence: 1.0, notes: "批量标注");
                Console.WriteLine($"  ✓ Body [{i}] {graphs[i].BodyName} → {value}");
            }
        
            Console.WriteLine($"\n✓ 成功标注 {bodyIds.Count} 个 body");
        
            // 显示统计信息
            TopologyLabeler.ShowStatistics();
        }
 
 
 [Command("find_category", Description = "查询当前零件的推荐类别（基于 WL 图核相似度匹配）", Parameters = "[可选的：迭代次数，默认 1] [可选的：返回数量，默认 5]", Group = "train")]
        static void FindCategory(string[] args)
        {
            if (Program.SwModel == null)
            {
                Console.WriteLine("错误：请先打开一个零件文档");
                return;
            }
     // 调试输出：显示 args 的实际内容
            Console.WriteLine($"[调试] FindCategory 收到 args.Length = {args.Length}");
            if (args.Length > 0)
            {
                Console.WriteLine($"[调试] args 内容：[{string.Join(", ", args)}]");
            }

            int iterations = 1;
            int topK = 5;

            if (args.Length > 0 && int.TryParse(args[0], out int iter))
            {
                iterations = iter;
            }

            if (args.Length > 1 && int.TryParse(args[1], out int k))
            {
                topK = k;
            }

            Console.WriteLine("=== 查询当前零件的推荐类别 ===\n");
            
            // 调用 TopologyLabeler.FindCategoriesByWL 方法
            TopologyLabeler.FindCategoriesByWL(Program.SwModel, wlIterations: iterations, topK: topK);
        }
        
        /// <summary>
        /// 显示 body 的现有标注信息
        /// </summary>
        static void DisplayBodyLabels(TopologyDatabase database, int bodyId)
        {
            var existingLabels = database.GetBodyLabels(bodyId);
            if (existingLabels.Count > 0)
            {
                Console.WriteLine("\n=== 现有标注 ===");
                foreach (var labelCategory in existingLabels)
                {
                    foreach (var labelData in labelCategory.Value)
                    {
                        Console.WriteLine($"  {labelCategory.Key}: {labelData.Item1} (置信度：{labelData.Item2}) - {labelData.Item3}");
                    }
                }
            }
        }

        /// <summary>
        /// 显示数据库中所有的标注类别
        /// </summary>
        [Command("view_categories", Description = "查看数据库中所有的标注类别", Parameters = "无", Group = "train")]
        static void ViewAllCategories(string[] args)
        {
            TopologyLabeler.Initialize();
            var database = TopologyLabeler.GetDatabase();
            
            Console.WriteLine("\n=== 所有标注类别 ===");
            var allCategories = database!.GetAllCategories();
            if (allCategories.Count > 0)
            {
                Console.WriteLine($"共 {allCategories.Count} 个类别:");
                foreach (var cat in allCategories)
                {
                    Console.WriteLine($"  - {cat}");
                }
            }
            else
            {
                Console.WriteLine("暂无任何标注类别");
            }
        }


        [Command("label_quick", Description = "快速标注当前零件（仅计算存储，不立即标注）", Parameters = "无", Group = "train")]
        static void QuickLabelPart(string[] args)
        {
            if (Program.SwModel == null)
            {
                Console.WriteLine("错误：请先打开一个零件文档");
                return;
            }

            TopologyLabelingExample.QuickLabel(Program.SwApp!, Program.SwModel);
        }

        [Command("view_parts", Description = "查看所有已标注的零件", Parameters = "无", Group = "train")]
        static void ViewAllParts(string[] args)
        {
            TopologyLabeler.ViewAllParts();
        }

        [Command("label_search", Description = "按标注类别搜索零件", Parameters = "[类别名称] [可选的值]", Group = "train")]
        static void SearchLabels(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("用法：label_search [类别名称] [可选的值]");
                Console.WriteLine("示例：label_search 结构类型 框架");
                return;
            }

            string category = args[0];
            string? value = args.Length > 1 ? args[1] : null;

            TopologyLabeler.SearchByCategory(category, value);
        }

        [Command("label_stats", Description = "显示数据库统计信息，打印一下数据库所有数据", Parameters = "无", Group = "train")]
        static void ShowLabelStats(string[] args)
        {
            TopologyLabeler.ShowStatistics();
        }

        [Command("label_batch", Description = "批量标注文件夹中的所有零件", Parameters = "[文件夹路径]", Group = "train")]
        static void BatchLabel(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("用法：label_batch [文件夹路径]");
                Console.WriteLine("示例：label_batch E:\\parts");
                return;
            }

            string folderPath = args[0];

            if (!System.IO.Directory.Exists(folderPath))
            {
                Console.WriteLine($"错误：文件夹不存在：{folderPath}");
                return;
            }

            if (Program.SwApp == null)
            {
                Console.WriteLine("错误：无法连接到 SolidWorks");
                return;
            }

            TopologyLabelingExample.BatchProcessExample(Program.SwApp, folderPath);
        }

        [Command("label_export", Description = "导出零件的 WL 结果为 JSON", Parameters = "[零件 ID]", Group = "train")]
        static void ExportWL(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("用法：label_export [零件 ID]");
                Console.WriteLine("提示：使用 view_parts 查看零件 ID");
                return;
            }

            if (int.TryParse(args[0], out int partId))
            {
                TopologyLabeler.ExportPartWL(partId);
            }
            else
            {
                Console.WriteLine("错误：零件 ID 必须是数字");
            }
        }

        [Command("label_delete", Description = "删除指定的标注 (删除零件 id 为 x 的数据)", Parameters = "[标注 ID]", Group = "train")]
        static void DeleteLabel(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("用法：label_delete [标注 ID]");
                Console.WriteLine("提示：使用 label_search 或 view_parts 查看标注 ID");
                return;
            }
        
            if (int.TryParse(args[0], out int labelId))
            {
                TopologyLabeler.DeleteLabel(labelId);
            }
            else
            {
                Console.WriteLine("错误：标注 ID 必须是数字");
            }
        }
        
        [Command("delete_part", Description = "删除零件 ID 为 x 的所有数据（包括 WL 结果和标注）", Parameters = "[零件 ID]", Group = "train")]
        static void DeletePart(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("用法：delete_part [零件 ID]");
                Console.WriteLine("提示：使用 view_parts 查看零件 ID");
                return;
            }
        
            if (int.TryParse(args[0], out int partId))
            {
                TopologyLabeler.DeletePartData(partId);
            }
            else
            {
                Console.WriteLine("错误：零件 ID 必须是数字");
            }
        }
    }
}