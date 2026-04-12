using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Text;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.IO;

namespace tools
{
    /// <summary>
    /// LLM 循环调用器（支持 Tool 调用模式）
    /// </summary>
    public class LlmLoopCaller
    {
        private readonly LlmService _llmService;
        private readonly CommandExecutor _commandExecutor;
        private StringWriter? _consoleOutputCapture;
        private TextWriter _originalConsoleOutput = null!;
        private readonly string _lastCommandFile;  // 本地记录文件路径
    
     
        private bool _requireConfirmation = true;
        
        // 记录最后执行的命令（从文件加载）
        private static string? _lastCommand = LoadLastCommandFromFile();
        
        // 特殊命令列表
        private readonly List<(string Name, string Description)> _specialCommands = new List<(string, string)> 
        { 
            ("quit", "退出交互式循环模式"),
            ("exit", "退出交互式循环模式"),
            ("clear", "清空对话历史"),
            ("mode", "切换命令执行模式（确认/自动）"),
            ("history", "查看对话历史"),
            ("last", "重复执行上一次执行的命令")
        };

        public LlmLoopCaller(
            Func<string>? getCommandsDescriptionFunc,
            Func<string, CommandInfo?> commandResolver,
            Func<SldWorks?> swAppResolver,
            Action<ModelDoc2?> swModelUpdater)
        {
            _llmService = new LlmService(getCommandsDescriptionFunc);
            _commandExecutor = new CommandExecutor(commandResolver, swAppResolver, swModelUpdater);

            // 初始化输出捕获
            _consoleOutputCapture = new StringWriter();
            _originalConsoleOutput = Console.Out;
            
            // 初始化本地记录文件路径
            string llmDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "llm");
            if (!Directory.Exists(llmDir))
            {
                Directory.CreateDirectory(llmDir);
            }
            _lastCommandFile = Path.Combine(llmDir, "last_command.txt");
            
            // 重定向 Console 输出（不再写入 runlog）
            Console.SetOut(_originalConsoleOutput);
        }

        /// <summary>
        /// 从文件加载最后执行的命令
        /// </summary>
        private static string? LoadLastCommandFromFile()
        {
            try
            {
                string llmDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "llm");
                string filePath = Path.Combine(llmDir, "last_command.txt");
                
                if (File.Exists(filePath))
                {
                    return File.ReadAllText(filePath, Encoding.UTF8).Trim();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[警告] 加载 last_command.txt 失败：{ex.Message}");
            }
            
            return null;
        }

        /// <summary>
        /// 保存最后执行的命令到文件
        /// </summary>
        private static void SaveLastCommandToFile(string command)
        {
            try
            {
                string llmDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "llm");
                if (!Directory.Exists(llmDir))
                {
                    Directory.CreateDirectory(llmDir);
                }
                
                string filePath = Path.Combine(llmDir, "last_command.txt");
                File.WriteAllText(filePath, command, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[警告] 保存 last_command.txt 失败：{ex.Message}");
            }
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
                
                // 为主命令名创建工具
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
                
                // 为每个别名创建工具
                if (cmd.Aliases != null)
                {
                    foreach (var alias in cmd.Aliases)
                    {
                        if (!string.IsNullOrEmpty(alias))
                        {
                            tools.Add(new ToolDefinition
                            {
                                Type = "function",
                                Function = new FunctionDefinition
                                {
                                    Name = $"execute_{alias}",
                                    Description = $"{cmd.Description ?? ""} (别名：{cmd.Name})",
                                    Parameters = parameters
                                }
                            });
                        }
                    }
                }
            }
            
            return tools;
        }

        /// <summary>
        /// 执行 Tool 调用
        /// </summary>
        private async Task<(string Result, string CapturedOutput)> ExecuteToolCallAsync(ToolCall toolCall)
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
                
                // 拦截 Console 输出到 tool 里面
                _consoleOutputCapture!.GetStringBuilder().Clear();
                Console.SetOut(_consoleOutputCapture);
                        
                // 等待用户确认
                if (_requireConfirmation)
                {
                    // 临时恢复原始 Console 输出，让用户能看到提示语
                    Console.SetOut(_originalConsoleOutput);
                    Console.Write("\n是否执行此命令？ (y/n/auto): ");
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
                        return ($"命令 '{functionName}' 已经被用户拒绝", "");
                    }
                                    
                    // 重新设置输出捕获，用于后续的命令执行
                    _consoleOutputCapture!.GetStringBuilder().Clear();
                    Console.SetOut(_consoleOutputCapture);
                }
                        
                Console.WriteLine($"\n>>> 正在执行命令：{fullCommand}...\n");
                
                try
                {
                    string result = await _commandExecutor.ExecuteCommandAsync(fullCommand);
                    
                    // 恢复 Console 输出
                    Console.SetOut(_originalConsoleOutput);
                    
                    // 获取捕获的输出
                    string capturedOutput = _consoleOutputCapture.ToString();
                    
                    Console.WriteLine($"\n[调试] ExecuteToolCallAsync - 执行结果：{result}");
                    
                    // 显示捕获的输出
                    if (!string.IsNullOrEmpty(capturedOutput))
                    {
                        Console.WriteLine("[捕获的输出]:");
                        Console.Write(capturedOutput);
                    }
                    
                    // 记录最后执行的命令（内存 + 文件）
                    _lastCommand = fullCommand;
                    SaveLastCommandToFile(fullCommand);
                    
                    return (result, capturedOutput);
                }
                catch
                {
                    // 发生异常时也要恢复 Console 输出
                    Console.SetOut(_originalConsoleOutput);
                    throw;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n[错误] ExecuteToolCallAsync 异常：{ex}");
                Console.WriteLine($"\n❌ 执行 Tool 调用失败：{ex.Message}");
                return ($"执行失败：{ex.Message}", "");
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
        /// 使用模糊匹配查找最接近的注册命令，返回完全匹配和模糊匹配结果
        /// </summary>
        private (string Command, double Score, bool IsExactMatch)? FindFuzzyCommand(string input, double threshold = 0.5)
        {
            string inputLower = input.ToLower().Trim();
            
            // 先检查是否是特殊命令（特殊命令不在此处处理）
            var specialCmd = FindFuzzySpecialCommand(input, threshold);
            if (specialCmd != null)
            {
                return null;
            }
            
            var allCommands = CommandRegistry.Instance.GetAllCommands();
            
            // 第一步：检查是否是 "命令名 + 参数" 的格式（优先处理）
            foreach (var cmd in allCommands.Values)
            {
                string cmdNameLower = cmd.Name.ToLower();
                
                // 检查输入是否以命令名开头（支持带空格参数）
                if (inputLower.StartsWith(cmdNameLower + " ") || inputLower == cmdNameLower)
                {
                    // 完全匹配命令名，带有参数
                    Console.WriteLine($"[调试] 命令 + 参数匹配：'{input}' -> '{cmd.Name}' (完全匹配命令名，带参数)");
                    return (cmd.Name, 1.0, true);
                }
                
                // 检查是否匹配别名
                if (cmd.Aliases != null)
                {
                    foreach (var alias in cmd.Aliases)
                    {
                        if (!string.IsNullOrEmpty(alias))
                        {
                            string aliasLower = alias.ToLower();
                            if (inputLower.StartsWith(aliasLower + " ") || inputLower == aliasLower)
                            {
                                // 完全匹配别名，带有参数
                                Console.WriteLine($"[调试] 命令 + 参数匹配：'{input}' -> '{cmd.Name}' (完全匹配别名 '{alias}',带参数)");
                                return (cmd.Name, 1.0, true);
                            }
                        }
                    }
                }
            }
            
            // 第二步：如果没有匹配到命令名，进行传统的模糊匹配（包括别名）
            var matches = new List<(string Command, double Score)>();
            
            foreach (var cmd in allCommands.Values)
            {
                // 计算与原名称的相似度
                double scoreName = CalculateSimilarity(inputLower, cmd.Name.ToLower());
                
                // 计算与别名的相似度（取最高分）
                double scoreAlias = 0.0;
                if (cmd.Aliases != null)
                {
                    foreach (var alias in cmd.Aliases)
                    {
                        if (!string.IsNullOrEmpty(alias))
                        {
                            double aliasScore = CalculateSimilarity(inputLower, alias.ToLower());
                            scoreAlias = Math.Max(scoreAlias, aliasScore);
                        }
                    }
                }
                
                // 计算与描述的相似度
                double scoreDesc = CalculateSimilarity(inputLower, (cmd.Description ?? "").ToLower());
                
                // 取最高分
                double score = Math.Max(scoreName, Math.Max(scoreAlias, scoreDesc));
                
                // 如果完全匹配（包括别名），给予最高分
                if (inputLower == cmd.Name.ToLower())
                {
                    score = 1.0;
                }
                else if (cmd.Aliases != null && cmd.Aliases.Any(a => a?.ToLower() == inputLower))
                {
                    // 如果是别名完全匹配，也给予最高分
                    score = 1.0;
                }
                
                if (score >= threshold)
                {
                    matches.Add((cmd.Name, score));
                }
            }
            
            // 返回得分最高的命令
            if (matches.Count > 0)
            {
                var bestMatch = matches.OrderByDescending(m => m.Score).First();
                bool isExactMatch = bestMatch.Score >= 1.0;
                Console.WriteLine($"[调试] 模糊匹配命令：'{input}' -> '{bestMatch.Command}' (相似度：{bestMatch.Score:F2}, {(isExactMatch ? "完全匹配" : "模糊匹配")})");
                return (bestMatch.Command, bestMatch.Score, isExactMatch);
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
            Console.WriteLine("输入 'history' 查看对话历史");
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
                  
                    continue;
                }
                        
                if (input.ToLower() == "mode")
                {
                    _requireConfirmation = !_requireConfirmation;
                    Console.WriteLine($"命令执行模式已切换为：{(_requireConfirmation ? "确认模式" : "自动模式")}\n");
                    continue;
                }
                
                if (input.ToLower() == "history")
                {
                    ViewHistory();
                    continue;
                }
                
                if (input.ToLower() == "last")
                {
                    if (string.IsNullOrEmpty(_lastCommand))
                    {
                        Console.WriteLine("\n还没有执行过任何命令喵~");
                    }
                    else
                    {
                        Console.WriteLine($"\n>>> 重复执行上一次命令：{_lastCommand}");
                        input = _lastCommand;  // 将 input 设置为上一次的命令，继续后续处理
                    }
                    continue;
                }
                
                if (input.ToLower() == "llm")
                {
                    await RunLlmModeAsync();
                    continue;
                }

                // 检查是否直接输入了命令名（完全匹配或模糊匹配）
                var matchedResult = FindFuzzyCommand(input);
                if (matchedResult.HasValue)
                {
                    string matchedCommand = matchedResult.Value.Command;
                    double matchScore = matchedResult.Value.Score;
                    bool isExactMatch = matchedResult.Value.IsExactMatch;
                    
                    // 完全匹配：直接调用
                    if (isExactMatch)
                    {
                        Console.WriteLine($"\n>>> 检测到完全匹配的命令：{matchedCommand}");
                        
                        try
                        {
                            string fullCommand = input;
                            
                            Console.WriteLine($"\n>>> 正在执行命令：{fullCommand}...\n");
                            
                            // 拦截 Console 输出
                            _consoleOutputCapture!.GetStringBuilder().Clear();
                            Console.SetOut(_consoleOutputCapture);
                            
                            try
                            {
                                // 临时恢复原始 Console 输出，让用户能看到命令内部的确认提示
                                Console.SetOut(_originalConsoleOutput);
                                
                                string result = await _commandExecutor.ExecuteCommandAsync(fullCommand);
                                
                                // 恢复 Console 输出
                                Console.SetOut(_originalConsoleOutput);
                                
                                // 获取捕获的输出
                                string capturedOutput = _consoleOutputCapture.ToString();
                                
                                Console.WriteLine($"\n[调试] ExecuteCommandAsync - 执行结果：{result}");
                                
                                // 显示捕获的输出
                                if (!string.IsNullOrEmpty(capturedOutput))
                                {
                                    Console.WriteLine("[捕获的输出]:");
                                    Console.Write(capturedOutput);
                                }
                                
                                // 将执行结果保存到记忆
                                SaveCommandResultToMemory(matchedCommand, result, capturedOutput);
                                
                                // 记录最后执行的命令（内存 + 文件）
                                _lastCommand = fullCommand;
                                SaveLastCommandToFile(fullCommand);
                            }
                            catch
                            {
                                // 发生异常时也要恢复 Console 输出
                                Console.SetOut(_originalConsoleOutput);
                                throw;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"\n[错误] 命令执行失败：{ex.Message}\n");
                        }
                        continue;
                    }
                    // 模糊匹配：发送给 LLM 处理
                    else
                    {
                        Console.WriteLine($"\n>>> 检测到模糊匹配的命令：'{input}' -> '{matchedCommand}' (相似度：{matchScore:F2})");
                        Console.WriteLine(">>> 将请求发送给 LLM 进一步确认...\n");
                    }
                }
        
                try
                {
                    // 使用 Tool 调用模式
                    var (response, toolCalls) = await _llmService.ChatWithToolsAsync(input, toolDefinitions);
                    
                    // 处理 Tool 调用
                    if (toolCalls != null && toolCalls.Count > 0)
                    {
                        Console.WriteLine($"\n>>> 检测到 {toolCalls.Count} 个命令调用请求");
                        
                        var results = new List<(string Result, string CapturedOutput)>();
                        foreach (var toolCall in toolCalls)
                        {
                            var (result, capturedOutput) = await ExecuteToolCallAsync(toolCall);
                            results.Add((result, capturedOutput));
                        }
                        
                        if (results.Count > 0)
                        {
                            Console.WriteLine($"\n>>> 所有命令执行完成");
                                                
                            // 将 Tool 调用结果保存到短期记忆（以 user 角色）
                            SaveToolResultsToMemory(toolCalls, results);
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
        /// 将 Tool 调用结果保存到短期记忆
        /// </summary>
        private void SaveToolResultsToMemory(List<ToolCall> toolCalls, List<(string Result, string CapturedOutput)> results)
        {
            try
            {
                // 构建 Tool 调用结果的系统消息
                var sb = new StringBuilder();
                sb.AppendLine("\n***Tool 调用执行结果:***");
                
                for (int i = 0; i < toolCalls.Count; i++)
                {
                    var toolCall = toolCalls[i];
                    var (result, capturedOutput) = results[i];
                    
                    // 解析函数名（去掉 execute_前缀）
                    string functionName = toolCall.Function.Name;
                    if (functionName.StartsWith("execute_"))
                    {
                        functionName = functionName.Substring(8);
                    }
                    
                    sb.AppendLine($"\n- 命令：{functionName}");
                    sb.AppendLine($"  返回：{result}");
                    if (!string.IsNullOrEmpty(capturedOutput))
                    {
                        sb.AppendLine($"  Console 输出:\n{capturedOutput.Trim()}");
                    }
                }
                
                // 将结果作为 user 消息保存（让 LLM 知道 Tool 执行的结果）
                var messages = LoadMessagesFromDisk();
                messages.Add(new ChatMessage 
                { 
                    Role = "user", 
                    Content = sb.ToString() 
                });
                
                SaveMessagesToDisk(messages);
                
                Console.WriteLine($"[调试] Tool 调用结果已保存到短期记忆（包含 Console 输出）");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[警告] 保存 Tool 调用结果失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 保存命令执行结果到记忆
        /// </summary>
        private void SaveCommandResultToMemory(string commandName, string result, string? consoleOutput = null)
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine("\n***命令执行结果:***");
                sb.AppendLine($"命令：{commandName}");
                sb.AppendLine($"结果：{result}");
                if (!string.IsNullOrEmpty(consoleOutput))
                {
                    sb.AppendLine($"Console 输出:\n{consoleOutput.Trim()}");
                }
                
                // 将结果作为 assistant 消息保存
                var messages = LoadMessagesFromDisk();
                messages.Add(new ChatMessage 
                { 
                    Role = "assistant", 
                    Content = sb.ToString() 
                });
                
                SaveMessagesToDisk(messages);
                
                Console.WriteLine($"[调试] 命令执行结果已保存到短期记忆{(consoleOutput != null ? "(包含 Console 输出)" : "")}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[警告] 保存命令执行结果失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 运行 LLM 对话模式
        /// </summary>
        private async Task RunLlmModeAsync()
        {
            Console.WriteLine("\n*** LLM 对话模式 ***");
            Console.WriteLine("在此模式下，AI 会以自然语言回答你的问题");
            Console.WriteLine("输入 'back' 返回命令执行模式\n");
            
            while (true)
            {
                var input = await GetUserInputAsync("你：");
                
                if (string.IsNullOrEmpty(input))
                {
                    continue;
                }
                
                input = input.Trim();
                
                if (input.ToLower() == "back")
                {
                    Console.WriteLine("\n返回命令执行模式\n");
                    break;
                }
                
                try
                {
                    var response = await _llmService.ChatAsync(input);
                    Console.WriteLine($"\n{response}\n");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"\n[错误] LLM 响应失败：{ex.Message}\n");
                }
            }
        }
        
        /// <summary>
        /// 从磁盘加载消息历史（复用 LlmService 的方法）
        /// </summary>
        private List<ChatMessage> LoadMessagesFromDisk()
        {
            try
            {
                string shotMemoryFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "llm", "shot_memory.json");
                if (File.Exists(shotMemoryFile))
                {
                    var json = File.ReadAllText(shotMemoryFile, System.Text.Encoding.UTF8);
                    var messages = JsonConvert.DeserializeObject<List<ChatMessage>>(json) ?? new List<ChatMessage>();
                    return messages.Where(m => m.Role != "system").ToList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"加载消息历史失败：{ex.Message}");
            }
            
            return new List<ChatMessage>();
        }
        
        /// <summary>
        /// 保存消息历史到磁盘（复用 LlmService 的方法）
        /// </summary>
        private void SaveMessagesToDisk(List<ChatMessage> messages)
        {
            try
            {
                string shotMemoryFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "llm", "shot_memory.json");
                var filteredMessages = messages.Where(m => m.Role != "system").ToList();
                
                // 如果消息超过 10 条，保留最近的 10 条（5 轮对话）
                if (filteredMessages.Count > 10)
                {
                    filteredMessages = filteredMessages.Skip(filteredMessages.Count - 10).ToList();
                }
                
                var json = JsonConvert.SerializeObject(filteredMessages, Formatting.Indented);
                File.WriteAllText(shotMemoryFile, json, System.Text.Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"保存消息历史失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 异步获取用户输入 (非阻塞)
        /// </summary>
        private async Task<string?> GetUserInputAsync(string prompt)
        {
            return await Task.Run(() =>
            {
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.Write(prompt);
                Console.ResetColor();
                  Console.ForegroundColor = ConsoleColor.Cyan;
                var input = Console.ReadLine();
                Console.ResetColor();
                return input;
            });
        }
        
        /// <summary>
        /// 查看对话历史（特殊命令，不通过 Tool 调用）
        /// </summary>
        private void ViewHistory()
        {
            Console.WriteLine("\n=== LLM 短期记忆 ===\n");
            
            // 查看短期记忆（shot_memory.json）
            string llmDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "llm");
            string shotMemoryFile = Path.Combine(llmDir, "shot_memory.json");
            
            if (File.Exists(shotMemoryFile))
            {
                try
                {
                    var json = File.ReadAllText(shotMemoryFile, Encoding.UTF8);
                    var messages = JsonConvert.DeserializeObject<List<dynamic>>(json);
                    if (messages != null && messages.Count > 0)
                    {
                        Console.WriteLine($"共 {messages.Count} 条消息:\n");
                        int i = 1;
                        foreach (var msg in messages)
                        {
                            string role = msg.role ?? "unknown";
                            string content = msg.content ?? "";
                            Console.WriteLine($"{i++}. [{role}] {content}");
                            Console.WriteLine();
                        }
                    }
                    else
                    {
                        Console.WriteLine("暂无对话历史");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"读取短期记忆失败：{ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("暂无短期记忆文件");
            }
        }

        /// <summary>
        /// 多路文本写入器 - 同时写入多个 TextWriter
        /// </summary>
        public class MultiTextWriter : TextWriter
        {
            private readonly List<TextWriter> _writers;

            public MultiTextWriter(params TextWriter[] writers)
            {
                _writers = new List<TextWriter>(writers ?? throw new ArgumentNullException(nameof(writers)));
            }

            public override Encoding Encoding => Encoding.UTF8;

            public override void Write(char value)
            {
                foreach (var writer in _writers)
                {
                    writer.Write(value);
                }
            }

            public override void Write(string? value)
            {
                foreach (var writer in _writers)
                {
                    writer.Write(value);
                }
            }

            public override void WriteLine(string? value)
            {
                foreach (var writer in _writers)
                {
                    writer.WriteLine(value);
                }
            }

            public override void Flush()
            {
                foreach (var writer in _writers)
                {
                    writer.Flush();
                }
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    foreach (var writer in _writers)
                    {
                        writer?.Dispose();
                    }
                }
                base.Dispose(disposing);
            }
        }




    }
}
