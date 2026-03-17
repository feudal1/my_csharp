using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace tools
{
    /// <summary>
    /// LLM 循环调用器
    /// 支持批量问题循环提问，自动保存每次对话结果
    /// </summary>
    public class LlmLoopCaller
    {
        private readonly LlmService _llmService;
        private readonly string _outputDir;
        private readonly string _loopHistoryFile;

        public LlmLoopCaller()
        {
            _llmService = new LlmService();

            // 初始化输出目录
            _outputDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "llm", "loop_output");
            if (!Directory.Exists(_outputDir))
            {
                Directory.CreateDirectory(_outputDir);
            }

            _loopHistoryFile = Path.Combine(_outputDir, "loop_history.txt");
        }



        /// <summary>
        /// 交互式循环调用（手动输入问题）
        /// </summary>
        public async Task InteractiveLoopAsync()
        {
            Console.WriteLine("\n进入交互式循环模式");
            Console.WriteLine("输入问题后按回车提交");
            Console.WriteLine("输入 'quit' 或 'exit' 退出");
            Console.WriteLine("输入 'clear' 清空对话历史\n");

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

                try
                {
                    var response = await _llmService.ChatAsync(input);
                 
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"\n[错误] 调用失败：{ex.Message}\n");
                }
            }
        }



    }
}
