using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

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
        private bool _requireConfirmation = true; // 默认需要用户确认

        public LlmLoopCaller()
        {
            _llmService = new LlmService();
            _commandExecutor = new CommandExecutor();

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
                    
                    string fullCommand = $"{commandName} {parameters}".Trim();
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
                            results.Add($"命令 '{commandName}' 已被用户拒绝");
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
        /// 交互式循环调用（手动输入问题）
        /// </summary>
        public async Task InteractiveLoopAsync()
        {
            Console.WriteLine("\n进入交互式循环模式");
            Console.WriteLine("输入问题后按回车提交") ;
            Debug.WriteLine("测试debug**************************");
            Console.WriteLine("输入 'quit' 或 'exit' 退出");
            Console.WriteLine("输入 'clear' 清空对话历史");
            Console.WriteLine("输入 'mode' 切换命令执行模式（确认/自动）");
            Console.WriteLine("AI 可以调用命令执行 SolidWorks 操作，格式：do_【命令名】参数");
            Console.WriteLine("例如：do_【export_dwg】C:\\path\\to\\file.sldprt");
            Console.WriteLine("当 AI 建议执行命令时，您可以选择:\n  y - 执行当前命令\n  n - 跳过当前命令\n  auto - 切换到自动模式，后续命令直接执行\n");

            while (true)
            {
                Console.Write("你：");
                var input = Console.ReadLine()?.Trim();

                if (string.IsNullOrEmpty(input))
                {
                    continue;
                }

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



    }
}
