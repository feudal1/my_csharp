using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace tools
{
    /// <summary>
    /// LLM 循环调用器（支持 Tool 调用模式）
    /// </summary>
    public class LlmLoopCaller
    {
        private readonly LlmService _llmService;
        private readonly CommandExecutor _commandExecutor;
        private readonly string _outputDir;
        private readonly string _loopHistoryFile;
        private bool _requireConfirmation = true;
        
        // 特殊命令列表
        private readonly List<(string Name, string Description)> _specialCommands = new List<(string, string)> 
        { 
            ("quit", "退出交互式循环模式"),
            ("exit", "退出交互式循环模式"),
            ("clear", "清空对话历史"),
            ("mode", "切换命令执行模式（确认/自动）")
        };

        public LlmLoopCaller(
            Func<string>? getCommandsDescriptionFunc,
            Func<string, CommandInfo?> commandResolver,
            Func<SldWorks?> swAppResolver,
            Action<ModelDoc2?> swModelUpdater)
        {
            _llmService = new LlmService(getCommandsDescriptionFunc);
            _commandExecutor = new CommandExecutor(commandResolver, swAppResolver, swModelUpdater);

            // 初始化输出目录
            _outputDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "llm", "loop_output");
            if (!Directory.Exists(_outputDir))
            {
                Directory.CreateDirectory(_outputDir);
            }

            _loopHistoryFile = Path.Combine(_outputDir, "loop_history.txt");
        }

        /// <summary>
        /// 从 CommandInfo 列表构建 Tool 定义
        /// </summary>
        private List<ToolDefinition> BuildToolDefinitions(Dictionary<string, CommandInfo> commands)
        {
            var tools = new List<ToolDefinition>();
            
            foreach (var cmd in commands.Values)
            {
                var parameters = new FunctionParameters();
                
                // 解析参数字符串，构建参数定义
                if (!string.IsNullOrEmpty(cmd.Parameters) && cmd.Parameters != "无")
                {
                    // 简单处理：将整个参数描述作为单个参数的说明
                    parameters.Properties["argument"] = new PropertyDefinition
                    {
                        Type = "string",
                        Description = cmd.Parameters
                    };
                    parameters.Required.Add("argument");
                }
                
                tools.Add(new ToolDefinition
                {
                    Type = "function",
                    Function = new FunctionDefinition
                    {
                        Name = $"execute_{cmd.Name}",
                        Description = cmd.Description ?? "",
                        Parameters = parameters
                    }
                });
            }
            
            return tools;
        }

        /// <summary>
        /// 执行 Tool 调用
        /// </summary>
        private async Task<string> ExecuteToolCallAsync(ToolCall toolCall)
        {
            try
            {
                Console.WriteLine($"\n[调试] ExecuteToolCallAsync: 开始处理 Tool 调用");
                Console.WriteLine($"[调试] ExecuteToolCallAsync: toolCall.Id = {toolCall.Id}");
                Console.WriteLine($"[调试] ExecuteToolCallAsync: toolCall.Function.Name = {toolCall.Function.Name}");
                Console.WriteLine($"[调试] ExecuteToolCallAsync: toolCall.Function.Arguments = {toolCall.Function.Arguments}");
                
                // 解析函数名（去掉 execute_前缀）
                string functionName = toolCall.Function.Name;
                if (functionName.StartsWith("execute_"))
                {
                    functionName = functionName.Substring(8);
                }
                
                Console.WriteLine($"[调试] ExecuteToolCallAsync: 解析后的函数名 = {functionName}");
                
                // 解析参数
                var arguments = JObject.Parse(toolCall.Function.Arguments);
                string argumentValue = arguments["argument"]?.ToString() ?? "";
                
                Console.WriteLine($"[调试] ExecuteToolCallAsync: argumentValue = '{argumentValue}'");
                
                string fullCommand = string.IsNullOrEmpty(argumentValue) 
                    ? functionName 
                    : $"{functionName} {argumentValue}";
                
                Console.WriteLine($"\n>>> Tool 调用：{functionName}");
                if (!string.IsNullOrEmpty(argumentValue))
                {
                    Console.WriteLine($"    参数：{argumentValue}");
                }
                
                // 等待用户确认
                if (_requireConfirmation)
                {
                    Console.Write("\n是否执行此命令？(y/n/auto): ");
                    var userInput = await GetUserInputAsync("");
                    userInput = userInput?.Trim().ToLower();
                    
                    if (userInput == "auto")
                    {
                        _requireConfirmation = false;
                        Console.WriteLine("已切换到自动模式，后续命令将直接执行");
                    }
                    else if (userInput != "y" && userInput != "yes")
                    {
                        Console.WriteLine("已跳过此命令");
                        return $"命令 '{functionName}' 已被用户拒绝";
                    }
                }
                
                Console.WriteLine($"\n>>> 正在执行命令：{fullCommand}...");
                Console.WriteLine($"[调试] ExecuteToolCallAsync - 函数名：{functionName}");
                Console.WriteLine($"[调试] ExecuteToolCallAsync - 参数值：{argumentValue}");
                Console.WriteLine($"[调试] ExecuteToolCallAsync - 完整命令：{fullCommand}");
                
                string result = await _commandExecutor.ExecuteCommandAsync(fullCommand);
                
                Console.WriteLine($"\n[调试] ExecuteToolCallAsync - 执行结果：{result}");
                Console.WriteLine($"执行结果：{result}");
                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n[错误] ExecuteToolCallAsync 异常：{ex}");
                Console.WriteLine($"\n❌ 执行 Tool 调用失败：{ex.Message}");
                return $"执行失败：{ex.Message}";
            }
        }


        /// <summary>
        /// 计算两个字符串的相似度（综合 Levenshtein 距离和字符集重叠度）
        /// </summary>
        private double CalculateSimilarity(string? query, string? text)
        {
            if (string.IsNullOrEmpty(query) || string.IsNullOrEmpty(text))
            {
                return 0.0;
            }

            string queryLower = query!.ToLower();
            string textLower = text!.ToLower();

            double ratio = CalculateLevenshteinRatio(queryLower, textLower);
            
            // 如果包含关键词，给予高分
            if (queryLower.Length > 0 && textLower.Contains(queryLower))
            {
                ratio = Math.Max(ratio, 0.9);
            }

            // 计算字符集重叠度
            var querySet = new HashSet<char>(queryLower);
            var textSet = new HashSet<char>(textLower);
            var overlap = querySet.Intersect(textSet).Count() / (double)Math.Max(querySet.Count, 1);

            // 综合评分：60% 编辑距离 + 40% 字符重叠
            return 0.6 * ratio + 0.4 * overlap;
        }

        /// <summary>
        /// 计算 Levenshtein 距离比率
        /// </summary>
        private double CalculateLevenshteinRatio(string s1, string s2)
        {
            if (string.IsNullOrEmpty(s1)) return string.IsNullOrEmpty(s2) ? 1.0 : 0.0;
            if (string.IsNullOrEmpty(s2)) return 0.0;

            int[,] matrix = new int[s1.Length + 1, s2.Length + 1];

            for (int i = 0; i <= s1.Length; i++)
                matrix[i, 0] = i;

            for (int j = 0; j <= s2.Length; j++)
                matrix[0, j] = j;

            for (int i = 1; i <= s1.Length; i++)
            {
                for (int j = 1; j <= s2.Length; j++)
                {
                    int cost = (s1[i - 1] == s2[j - 1]) ? 0 : 1;
                    matrix[i, j] = Math.Min(
                        Math.Min(
                            matrix[i - 1, j] + 1,
                            matrix[i, j - 1] + 1
                        ),
                        matrix[i - 1, j - 1] + cost
                    );
                }
            }

            int distance = matrix[s1.Length, s2.Length];
            int maxLen = Math.Max(s1.Length, s2.Length);
            return maxLen > 0 ? 1.0 - (distance / (double)maxLen) : 1.0;
        }

        /// <summary>
        /// 使用模糊匹配查找最接近的特殊命令
        /// </summary>
        private string? FindFuzzySpecialCommand(string input, double threshold = 0.6)
        {
            string inputLower = input.ToLower().Trim();
            
            var matches = new List<(string Command, double Score)>();
            
            foreach (var cmd in _specialCommands)
            {
                double score = CalculateSimilarity(inputLower, cmd.Name);
                
                if (score >= threshold)
                {
                    matches.Add((cmd.Name, score));
                }
            }
            
            // 返回得分最高的命令
            if (matches.Count > 0)
            {
                return matches.OrderByDescending(m => m.Score).First().Command;
            }
            
            return null;
        }

        /// <summary>
        /// 交互式循环调用（使用 Tool 调用模式）
        /// </summary>
        public async Task InteractiveLoopAsync()
        {
            Console.WriteLine("\n进入交互式循环模式（Tool 调用模式）");
            Console.WriteLine("输入问题后按回车提交");
            Debug.WriteLine("测试 debug**************************");
            Console.WriteLine("输入 'quit' 或 'exit' 退出");
            Console.WriteLine("输入 'clear' 清空对话历史");
            Console.WriteLine("输入 'mode' 切换命令执行模式（确认/自动）");
            Console.WriteLine("AI 会自动识别并调用合适的命令\n");
            
            // 获取所有可用命令并构建 Tool 定义
            var allCommands = CommandRegistry.Instance.GetAllCommands();
            var toolDefinitions = BuildToolDefinitions(allCommands);
            Console.WriteLine($"已加载 {toolDefinitions.Count} 个可用命令\n");
        
            while (true)
            {
                // 获取用户输入
                var input = await GetUserInputAsync("你：");
                                
                if (string.IsNullOrEmpty(input))
                {
                    continue;
                }
                                
                input = input!.Trim();
        
                if (input.ToLower() == "quit" || input.ToLower() == "exit")
                {
                    Console.WriteLine("退出交互式循环模式");
                    break;
                }
        
                if (input.ToLower() == "clear")
                {
                    _llmService.ClearHistory();
                    Console.WriteLine("对话历史已清空\n");
                    continue;
                }
                        
                if (input.ToLower() == "mode")
                {
                    _requireConfirmation = !_requireConfirmation;
                    Console.WriteLine($"命令执行模式已切换为：{(_requireConfirmation ? "确认模式" : "自动模式")}\n");
                    continue;
                }
        
                try
                {
                    // 使用 Tool 调用模式
                    var (response, toolCalls) = await _llmService.ChatWithToolsAsync(input, toolDefinitions);
                    
                    // 处理 Tool 调用
                    if (toolCalls != null && toolCalls.Count > 0)
                    {
                        Console.WriteLine($"\n>>> 检测到 {toolCalls.Count} 个命令调用请求");
                        
                        var results = new List<string>();
                        foreach (var toolCall in toolCalls)
                        {
                            var result = await ExecuteToolCallAsync(toolCall);
                            results.Add(result);
                        }
                        
                        if (results.Count > 0)
                        {
                            Console.WriteLine($"\n>>> 所有命令执行完成");
                        }
                    }
                    else if (!string.IsNullOrEmpty(response))
                    {
                        // 普通文本回复
                        Console.WriteLine($"\n{response}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"\n[错误] 调用失败：{ex.Message}\n");
                }
            }
        }
        
        /// <summary>
        /// 异步获取用户输入（非阻塞）
        /// </summary>
        private async Task<string?> GetUserInputAsync(string prompt)
        {
            // 直接使用控制台输入
            return await Task.Run(() =>
            {
                Console.Write(prompt);
                return Console.ReadLine();
            });
        }




    }
}
