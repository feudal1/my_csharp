using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using tools;

namespace tools
{
    /// <summary>
    /// LLM 标注验证器 - 使用 AI 检查标注结果是否正确
    /// </summary>
    public class LabelValidator
    {
        private readonly LlmService _llmService;
        
        public LabelValidator(Func<string>? getCommandsDescriptionFunc = null)
        {
            _llmService = new LlmService(getCommandsDescriptionFunc);
        }

        /// <summary>
        /// 核心方法：根据零件的 WL 拓扑图特征和标注值，让 AI 判断标注是否正确
        /// </summary>
        /// <param name="partName">零件名称</param>
        /// <param name="labelValue">标注值</param>
        /// <param name="wlFrequencies">WL 频率数据（最后一轮迭代）</param>
        /// <param name="drawingInfo">图纸信息（可选）</param>
        /// <returns>true 表示标注正确，false 表示标注错误</returns>
        public async Task<bool> ValidateLabelAsync(string partName, string labelValue, List<Dictionary<string, int>> wlFrequencies, string? drawingInfo = null)
        {
            try
            {
                // 使用 WL 特征查找数据库中相似的标注
                var database = TopologyLabeler.GetDatabase();
                if (database == null)
                {
                    Console.WriteLine("[警告] 数据库未初始化，仅基于零件名称验证");
                    return await ValidateByLLMOnly(partName, labelValue, drawingInfo);
                }
                
                // 调用现有的 WL 匹配方法查找相似标注
                var similarBodies = database.FindBodiesByWLTags(wlFrequencies, topK: 5, minSimilarity: 0.3);
                
                // 构建验证 prompt
                var promptBuilder = new StringBuilder();
                promptBuilder.AppendLine("请作为专业的 CAD 工程师，帮我判断以下零件的标注是否正确。");
                promptBuilder.AppendLine("\n=== 零件信息 ===");
                promptBuilder.AppendLine($"零件名称：{partName}");
                promptBuilder.AppendLine($"待验证标注值：{labelValue}");
                
                if (similarBodies.Count > 0)
                {
                    promptBuilder.AppendLine("\n=== 数据库中相似的已标注零件（基于 WL 拓扑特征匹配）===");
                    for (int i = 0; i < Math.Min(similarBodies.Count, 5); i++)
                    {
                        var (partId, simPartName, bodyId, bodyName, category, simLabelValue, similarity, confidence, notes) = similarBodies[i];
                        promptBuilder.AppendLine($"{i + 1}. {simPartName}+{bodyName}");
                        promptBuilder.AppendLine($"   标注：{category} = {simLabelValue}");
                        promptBuilder.AppendLine($"   相似度：{similarity:F3} ({similarity * 100:F1}%)");
                        promptBuilder.AppendLine($"   置信度：{confidence:F2}");
                    }
                }
                else
                {
                    promptBuilder.AppendLine("\n=== 数据库中没有找到相似的已标注零件 ===");
                }
                
                if (!string.IsNullOrEmpty(drawingInfo))
                {
                    promptBuilder.AppendLine($"\n=== 图纸信息 ===\n{drawingInfo}");
                }
                
                promptBuilder.AppendLine("\n=== 判断要求 ===");
                promptBuilder.AppendLine("1. 分析零件名称、WL 拓扑特征匹配的相似零件标注");
                promptBuilder.AppendLine("2. 对比待验证标注与相似零件的标注是否一致或合理");
                promptBuilder.AppendLine("3. 如果相似零件都标注为同一类型，而待验证标注不同，则可能错误");
                promptBuilder.AppendLine("4. 只返回 'true' 或 'false'，不要有其他内容");
                promptBuilder.AppendLine("5. true 表示标注正确，false 表示标注错误");
                
                string prompt = promptBuilder.ToString();
                
                Console.WriteLine($"\n[调试] 正在请求 LLM 验证标注...");
                Console.WriteLine($"[调试] 零件名称：{partName}");
                Console.WriteLine($"[调试] 标注值：{labelValue}");
                Console.WriteLine($"[调试] 找到 {similarBodies.Count} 个相似零件");
                
                // 调用 LLM
                string response = await _llmService.ChatAsync(prompt);
                
                // 解析响应，提取 true/false
                bool result = ParseBoolResponse(response);
                
                Console.WriteLine($"[调试] LLM 原始响应：{response.Trim()}");
                Console.WriteLine($"[调试] 解析结果：{result}");
                
                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n[错误] 标注验证失败：{ex.Message}");
                // 发生异常时返回 false，表示无法验证
                return false;
            }
        }

        /// <summary>
        /// 仅基于 LLM 验证（不使用 WL 特征）
        /// </summary>
        private async Task<bool> ValidateByLLMOnly(string partName, string labelValue, string? drawingInfo = null)
        {
            var promptBuilder = new StringBuilder();
            promptBuilder.AppendLine("请作为专业的 CAD 工程师，帮我判断以下零件的标注是否正确。");
            promptBuilder.AppendLine("\n=== 零件信息 ===");
            promptBuilder.AppendLine($"零件名称：{partName}");
            promptBuilder.AppendLine($"标注值：{labelValue}");
            
            if (!string.IsNullOrEmpty(drawingInfo))
            {
                promptBuilder.AppendLine($"\n=== 图纸信息 ===\n{drawingInfo}");
            }
            
            promptBuilder.AppendLine("\n=== 判断要求 ===");
            promptBuilder.AppendLine("1. 分析零件名称和图纸特征");
            promptBuilder.AppendLine("2. 判断该标注值是否符合零件的实际类型");
            promptBuilder.AppendLine("3. 只返回 'true' 或 'false'，不要有其他内容");
            promptBuilder.AppendLine("4. true 表示标注正确，false 表示标注错误");
            
            string prompt = promptBuilder.ToString();
            string response = await _llmService.ChatAsync(prompt);
            return ParseBoolResponse(response);
        }

        /// <summary>
        /// 解析 LLM 响应，提取布尔值
        /// </summary>
        private bool ParseBoolResponse(string response)
        {
            if (string.IsNullOrEmpty(response))
            {
                return false;
            }
            
            string lowerResponse = response.ToLower().Trim();
            
            // 直接匹配 true/false
            if (lowerResponse.Contains("true"))
            {
                return true;
            }
            if (lowerResponse.Contains("false"))
            {
                return false;
            }
            
            // 匹配中文
            if (lowerResponse.Contains("正确") || lowerResponse.Contains("是") || lowerResponse.Contains("对"))
            {
                return true;
            }
            if (lowerResponse.Contains("错误") || lowerResponse.Contains("否") || lowerResponse.Contains("不对"))
            {
                return false;
            }
            
            // 默认返回 false
            return false;
        }

        /// <summary>
        /// 获取当前零件的所有标注信息
        /// </summary>
        private string GetCurrentPartLabels()
        {
            if (Program.SwModel == null)
            {
                return "未打开任何零件";
            }
            
            var database = TopologyLabeler.GetDatabase();
            if (database == null)
            {
                return "数据库未初始化";
            }
            
            string partName = Program.SwModel.GetTitle();
            string fullPath = Program.SwModel.GetPathName();
            
            // 查询该零件的所有 body 标注
            var allBodies = database.GetAllBodiesWithLabels();
            var partBodies = allBodies.Where(b => b.PartName == partName).ToList();
            
            if (partBodies.Count == 0)
            {
                return $"零件 '{partName}' 没有标注记录";
            }
            
            var sb = new StringBuilder();
            sb.AppendLine($"零件：{partName}");
            sb.AppendLine($"路径：{fullPath}");
            sb.AppendLine($"\n共 {partBodies.Count} 个 Body:");
            
            foreach (var (partId, pName, bodyId, bodyName, labels) in partBodies)
            {
                sb.AppendLine($"\n- {bodyName}");
                if (labels.Count > 0)
                {
                    foreach (var label in labels)
                    {
                        sb.AppendLine($"  {label.Key}: {label.Value}");
                    }
                }
                else
                {
                    sb.AppendLine("  (无标注)");
                }
            }
            
            return sb.ToString();
        }

        /// <summary>
        /// 通过 Tool 方式验证当前零件的标注
        /// 这是一个可以被 LLM 调用的工具方法
        /// </summary>
        /// <param name="args">参数数组，包含：[零件名称, 标注值, 图纸信息(可选)]</param>
        public static async Task ValidateLabelTool(string[] args)
        {
            try
            {
                if (args.Length < 2)
                {
                    Console.WriteLine("用法：validate_label [零件名称] [标注值] [图纸信息(可选)]");
                    Console.WriteLine("示例：validate_label 连接板 结构件");
                    Console.WriteLine("示例：validate_label 法兰盘 管件 圆形法兰，带螺栓孔");
                    return;
                }
                
                string partName = args[0];
                string labelValue = args[1];
                string? drawingInfo = args.Length > 2 ? args[2] : null;
                
                Console.WriteLine($"\n=== LLM 标注验证 ===");
                Console.WriteLine($"零件名称：{partName}");
                Console.WriteLine($"标注值：{labelValue}");
                if (!string.IsNullOrEmpty(drawingInfo))
                {
                    Console.WriteLine($"图纸信息：{drawingInfo}");
                }
                
                // 获取当前零件的 WL 拓扑图特征
                Console.WriteLine("\n[调试] 正在计算 WL 拓扑图特征...");
                var graphs = FaceGraphBuilder.BuildGraphs(Program.SwModel);
                if (graphs.Count == 0)
                {
                    Console.WriteLine("无法构建拓扑图");
                    return;
                }
                
                // 合并所有 Body 的 WL 频率
                var combinedFrequencies = new List<Dictionary<string, int>>();
                foreach (var graph in graphs)
                {
                    var wlFreq = WLGraphKernel.PerformWLIterations(graph, 1);
                    if (combinedFrequencies.Count == 0)
                    {
                        combinedFrequencies = wlFreq;
                    }
                    else
                    {
                        for (int i = 0; i < wlFreq.Count; i++)
                        {
                            foreach (var kvp in wlFreq[i])
                            {
                                if (!combinedFrequencies[i].ContainsKey(kvp.Key))
                                    combinedFrequencies[i][kvp.Key] = 0;
                                combinedFrequencies[i][kvp.Key] += kvp.Value;
                            }
                        }
                    }
                }
                Console.WriteLine($"[调试] WL 特征计算完成（{combinedFrequencies.Count} 轮迭代）");
                
                var validator = new LabelValidator();
                bool isValid = await validator.ValidateLabelAsync(partName, labelValue, combinedFrequencies, drawingInfo);
                
                Console.WriteLine($"\n{'=',60}");
                Console.WriteLine($"验证结果：{(isValid ? "✓ 标注正确" : "✗ 标注错误")}");
                Console.WriteLine($"{'=',60}\n");
                
                if (!isValid)
                {
                    Console.WriteLine("建议：");
                    Console.WriteLine("1. 检查零件名称是否准确");
                    Console.WriteLine("2. 查看图纸确认零件实际类型");
                    Console.WriteLine("3. 参考历史标注数据重新标注");
                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n[错误] 验证过程出错：{ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        /// <summary>
        /// 批量验证当前零件的所有标注
        /// </summary>
        public static async Task ValidateAllLabelsTool(string[] args)
        {
            try
            {
                if (Program.SwModel == null)
                {
                    Console.WriteLine("错误：请先打开一个零件文档");
                    return;
                }
                
                Console.WriteLine($"\n=== 批量验证所有标注 ===\n");
                
                var database = TopologyLabeler.GetDatabase();
                if (database == null)
                {
                    Console.WriteLine("错误：数据库未初始化");
                    return;
                }
                
                string partName = Program.SwModel.GetTitle();
                string fullPath = Program.SwModel.GetPathName();
                
                // 构建当前零件的拓扑图
                Console.WriteLine("[调试] 正在构建拓扑图...");
                var graphs = FaceGraphBuilder.BuildGraphs(Program.SwModel);
                if (graphs.Count == 0)
                {
                    Console.WriteLine("无法构建拓扑图");
                    return;
                }
                Console.WriteLine($"[调试] 构建了 {graphs.Count} 个 Body 的拓扑图");
                
                // 获取该零件的所有 body 标注
                var allBodies = database.GetAllBodiesWithLabels();
                var partBodies = allBodies.Where(b => b.PartName == partName).ToList();
                
                if (partBodies.Count == 0)
                {
                    Console.WriteLine($"零件 '{partName}' 没有标注记录");
                    return;
                }
                
                var validator = new LabelValidator();
                int correctCount = 0;
                int incorrectCount = 0;
                int skippedCount = 0;
                
                foreach (var (partId, pName, bodyId, bodyName, labels) in partBodies)
                {
                    if (labels.Count == 0)
                    {
                        skippedCount++;
                        continue;
                    }
                    
                    // 取第一个标注进行验证（通常只有一个标注类别）
                    var firstLabel = labels.First();
                    string category = firstLabel.Key;
                    string labelValue = firstLabel.Value;
                    
                    Console.WriteLine($"\n验证 Body: {bodyName}");
                    Console.WriteLine($"  标注：{category} = {labelValue}");
                    
                    // 获取该 Body 的 WL 特征
                    var bodyGraph = graphs.FirstOrDefault(g => g.FullBodyName == bodyName);
                    List<Dictionary<string, int>> bodyWLFreq = new List<Dictionary<string, int>>();
                    if (bodyGraph != null)
                    {
                        bodyWLFreq = WLGraphKernel.PerformWLIterations(bodyGraph, 1);
                    }
                    
                    // 调用 LLM 验证
                    bool isValid = await validator.ValidateLabelAsync(bodyName, labelValue, bodyWLFreq);
                    
                    if (isValid)
                    {
                        Console.WriteLine($"  结果：✓ 正确");
                        correctCount++;
                    }
                    else
                    {
                        Console.WriteLine($"  结果：✗ 错误");
                        incorrectCount++;
                    }
                }
                
                Console.WriteLine($"\n{'=',60}");
                Console.WriteLine($"验证完成统计：");
                Console.WriteLine($"  正确：{correctCount}");
                Console.WriteLine($"  错误：{incorrectCount}");
                Console.WriteLine($"  跳过（无标注）：{skippedCount}");
                Console.WriteLine($"{'=',60}\n");
                
                if (incorrectCount > 0)
                {
                    Console.WriteLine("提示：发现错误的标注，建议重新检查并修正");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n[错误] 批量验证过程出错：{ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
