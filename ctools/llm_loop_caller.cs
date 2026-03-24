using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;

namespace tools
{
    /// <summary>
    /// LLM 循环调用器
    /// 支持批量问题循环提问，自动保存每次对话结果
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
        }; // 默认需要用户确认

        public LlmLoopCaller(
            Func<string>? getCommandsDescriptionFunc,
            Func<string, CommandInfo?> commandResolver,
            Func<SldWorks?> swAppResolver)
        {
            _llmService = new LlmService(getCommandsDescriptionFunc);
            _commandExecutor = new CommandExecutor(commandResolver, swAppResolver);

            // 初始化输出目录
            _outputDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "llm", "loop_output");
            if (!Directory.Exists(_outputDir))
            {
                Directory.CreateDirectory(_outputDir);
            }

            _loopHistoryFile = Path.Combine(_outputDir, "loop_history.txt");
        }


        /// <summary>
        /// 从 AI 响应中提取并执行命令（需要用户确认）
        /// </summary>
        private async Task<string> ProcessAIResponseAsync(string response)
        {
            // 匹配 do_【命令名】参数 模式（更严格的格式）
            var pattern = @"do_【([^】]+)】\s*(.*)";
            var matches = Regex.Matches(response, pattern);

            if (matches.Count > 0)
            {
                Console.WriteLine("\n>>> 检测到命令调用请求");
                
                var results = new List<string>();
                foreach (Match match in matches)
                {
                    string commandName = match.Groups[1].Value.Trim();
                    string parameters = match.Groups[2].Value.Trim();
                    
                    string fullCommand = string.IsNullOrEmpty(parameters) 
                        ? commandName 
                        : $"{commandName} {parameters}";
                    Console.WriteLine($"\n建议执行的命令：{fullCommand}");
                    
                    // 等待用户确认
                    if (_requireConfirmation)
                    {
                        Console.Write("\n是否执行此命令？(y/n/auto): ");
                        var userInput = Console.ReadLine()?.Trim().ToLower();
                        
                        if (userInput == "auto")
                        {
                            _requireConfirmation = false;
                            Console.WriteLine("已切换到自动模式，后续命令将直接执行");
                        }
                        else if (userInput != "y" && userInput != "yes")
                        {
                            Console.WriteLine("已跳过此命令");
                            results.Add($"命令 '{commandName} {parameters}' 已被用户拒绝");
                            continue;
                        }
                    }
                    
                    Console.WriteLine($"\n>>> 正在执行命令：{fullCommand}...");
                    
                    string result = await _commandExecutor.ExecuteCommandAsync(fullCommand);
                    results.Add(result);
                    
                    Console.WriteLine($"执行结果：{result}");
                }

                return string.Join("\n", results);
            }

            return "";
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
        /// 交互式循环调用（手动输入问题）
        /// </summary>
        public async Task InteractiveLoopAsync()
        {
            Console.WriteLine("\n进入交互式循环模式");
            Console.WriteLine("输入问题后按回车提交") ;
            Debug.WriteLine("测试 debug**************************");
            Console.WriteLine("输入 'quit' 或 'exit' 退出");
            Console.WriteLine("输入 'clear' 清空对话历史");
            Console.WriteLine("输入 'mode' 切换命令执行模式（确认/自动）");
            Console.WriteLine("AI 可以调用命令执行 SolidWorks 操作，格式：do_【命令名】参数");
            Console.WriteLine("例如：do_【export_dwg】C:\\path\\to\\file.sldprt");
            Console.WriteLine("当 AI 建议执行命令时，您可以选择:\n  y - 执行当前命令\n  n - 跳过当前命令\n  auto - 切换到自动模式，后续命令直接执行\n");
        
            while (true)
            {
                // 使用非阻塞方式获取用户输入
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
                
                // 检查是否有拼写错误（模糊匹配特殊命令）
                var fuzzyCmd = FindFuzzySpecialCommand(input, threshold: 0.6);
                if (fuzzyCmd != null)
                {
                    Console.WriteLine($"\n⚠️  检测到您可能想输入 '{fuzzyCmd}'，是否执行？(y/n): ");
                    var confirm = Console.ReadLine()?.Trim().ToLower();
                    
                    if (confirm == "y" || confirm == "yes")
                    {
                        if (fuzzyCmd == "quit" || fuzzyCmd == "exit")
                        {
                            Console.WriteLine("退出交互式循环模式");
                            break;
                        }
                        
                        if (fuzzyCmd == "clear")
                        {
                            _llmService.ClearHistory();
                            Console.WriteLine("对话历史已清空\n");
                            continue;
                        }
                        
                        if (fuzzyCmd == "mode")
                        {
                            _requireConfirmation = !_requireConfirmation;
                            Console.WriteLine($"命令执行模式已切换为：{(_requireConfirmation ? "确认模式" : "自动模式")}\n");
                            continue;
                        }
                    }
                    else
                    {
                        Console.WriteLine("已取消操作\n");
                    }
                }
        
                try
                {
                    var response = await _llmService.ChatAsync(input);
                            
                    // 检查并执行 AI 响应中的命令
                    var commandResult = await ProcessAIResponseAsync(response);
                            
                    if (!string.IsNullOrEmpty(commandResult))
                    {
                        Console.WriteLine($"\n>>> 命令执行完成，结果：{commandResult}");
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
