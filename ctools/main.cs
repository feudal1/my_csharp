using System;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

using cad_tools;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;
using System.Text.Json;
using System.Diagnostics;
using System.Text;

namespace tools
{
    // 性能监控装饰器属性
    [System.AttributeUsage(System.AttributeTargets.Method)]
    class ProfiledAttribute : System.Attribute
    {
        public string? Description { get; set; }
    }
    
    // 命令类型枚举 (公开)
    public enum CommandType
    {
        Sync,      // 同步命令
        Async      // 异步命令
    }
    public partial class Program
    {
        static Dictionary<string, Func<string[], Task>>? asyncCommands;
        static SldWorks? swApp;
        static ModelDoc2? swModel;
        
        // 添加公共静态属性，允许外部访问命令字典和 SolidWorks 实例
        public static Dictionary<string, CommandInfo>? Commands => commandInfos;
        public static SldWorks? SwApp => swApp;
        public static ModelDoc2? SwModel => swModel;
        
        // 命令信息类 (公开)
        public class CommandInfo
        {
          public string Name { get; set; } = "";
          public string? Description { get; set; }
          public string? Parameters { get; set; }
          public string? Group { get; set; }  // 命令所属组
          public Func<string[], Task> AsyncAction { get; set; } = null!;
          public CommandType CommandType { get; set; } = CommandType.Sync;  // 命令类型标记
        }
        
        static Dictionary<string, CommandInfo>? commandInfos;
        
        [STAThread]
        static void Main(string[] args)
        {
            // 设置控制台编码为 UTF-8，确保中文字符正常显示
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            
            RegisterCommands();
            GenerateCommandsDescriptionFile();
            
            if (args.Length==0)
            {
                LlmLoopCaller  loopCaller= new LlmLoopCaller();
           
                var task = Task.Run(() =>      loopCaller.InteractiveLoopAsync());
                task.GetAwaiter().GetResult();
            }

            // 仅导出命令信息时，不需要连接 SolidWorks（否则在未启动 SW 时会直接失败）
            if (args.Length > 0)
            {
                string command = args[0];
                
             
                
                if (command == "--search" || command == "search" || command == "s")
                {
                    string keyword = args.Length > 1 ? args[1] : "";
                    double threshold = 0.3;
                    int? topK = null;
                    
                    if (args.Length > 2 && double.TryParse(args[2], out double t))
                    {
                        threshold = t;
                    }
                    
                    if (args.Length > 3 && int.TryParse(args[3], out int k))
                    {
                        topK = k;
                    }
                    
                    SearchCommands(keyword, threshold, topK);
                    return;
                }
            }

            // 其余命令才需要连接 SolidWorks
            swApp = Connect.run();
            if (swApp == null)
            {
                Console.WriteLine("错误：无法连接到 SolidWorks 应用程序。");
                ShowHelp();
                return;
            }

            swModel = (ModelDoc2)swApp.ActiveDoc;

            if (args.Length > 0)
            {
                string command = args[0];

                if (asyncCommands != null && asyncCommands.ContainsKey(command))
                {
                    Console.WriteLine($">>> 正在执行命令：{command}...\n");
                    
                    // 根据命令类型选择执行方式
                    var cmdInfo = commandInfos![command];
                    
                    if (cmdInfo.CommandType == CommandType.Async)
                    {
                        // 异步命令：在新线程上运行，避免阻塞控制台输出
                        // 使用 Task.Run 启动异步任务，然后在当前线程等待
                        // 这样 Console.Out.Flush() 可以正常工作
                        var task = Task.Run(() => asyncCommands[command](args));
                        task.GetAwaiter().GetResult();
                    }
                    else
                    {
                        // 同步命令：直接调用
                        asyncCommands[command](args).GetAwaiter().GetResult();
                    }
                    
                    Console.WriteLine("\n>>> 命令执行结束。");
                }
                else
                {
                    Console.WriteLine($"没这命令：{command}");
                    ShowHelp();
                }
            }
    
        }
        
        /// <summary>
        /// 生成命令描述文件供 AI 使用
        /// </summary>
        static void GenerateCommandsDescriptionFile()
        {
            try
            {
                string llmDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "llm");
                if (!Directory.Exists(llmDir))
                {
                    Directory.CreateDirectory(llmDir);
                }
                
                string commandsFile = Path.Combine(llmDir, "commands_description.txt");
                
                var sb = new StringBuilder();
                sb.AppendLine("=== SolidWorks 自动化命令列表 ===");
                sb.AppendLine($"生成时间：{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                sb.AppendLine("\n***命令格式说明:***");
                sb.AppendLine("1. 无参数命令：直接输入 do_【命令名】");
                sb.AppendLine("2. 有参数命令：do_【命令名】参数值");
                sb.AppendLine("3. 执行前需用户确认 (y/n/auto)\n");
                
                if (commandInfos != null)
                {
                    foreach (var cmd in commandInfos.Values.OrderBy(k => k.Name))
                    {
                        sb.AppendLine($"\n【{cmd.Group ?? "默认"}】{cmd.Name} {(cmd.CommandType == CommandType.Async ? "(异步)" : "")}");
                        if (!string.IsNullOrEmpty(cmd.Description))
                        {
                            sb.AppendLine($"    说明：{cmd.Description}");
                        }
                        
                        // 明确标识参数
                        if (string.IsNullOrEmpty(cmd.Parameters) || cmd.Parameters == "无")
                        {
                            sb.AppendLine($"    参数：无");
                            sb.AppendLine($"    示例：do_【{cmd.Name}】");
                        }
                        else
                        {
                            sb.AppendLine($"    参数：{cmd.Parameters}");
                            sb.AppendLine($"    示例：do_【{cmd.Name}】<参数值>");
                        }
                    }
                }
                
                File.WriteAllText(commandsFile, sb.ToString(), Encoding.UTF8);
                Console.WriteLine($"\n✓ 命令描述文件已生成：{commandsFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n生成命令描述文件失败：{ex.Message}");
            }
        }
        

       static void ShowHelp()
        {
           Console.WriteLine("\n可用命令:");
            if (commandInfos != null)
            {
                foreach (var cmd in commandInfos.Values.OrderBy(k => k.Name))
                {
                   Console.WriteLine($"\n【{cmd.Group}】 {cmd.Name} {(cmd.CommandType == CommandType.Async ? "(异步)" : "")}");
                    if (!string.IsNullOrEmpty(cmd.Description))
                    {
                       Console.WriteLine($"    说明：{cmd.Description}");
                    }
                    if (!string.IsNullOrEmpty(cmd.Parameters))
                    {
                       Console.WriteLine($"    参数：{cmd.Parameters}");
                    }
                }
            }
           Console.WriteLine("\n使用方法：<命令> [参数...]");
           Console.WriteLine("查看帮助：<命令> -h 或 <命令> --help");
        }
        
      static void RegisterCommands()
        {
            commandInfos = new Dictionary<string, CommandInfo>();
            
            var methods = typeof(Program).GetMethods(BindingFlags.NonPublic | BindingFlags.Static)
                .Where(m => m.GetCustomAttribute<CommandAttribute>() != null);
            
            foreach (var method in methods)
            {
                var attr = method.GetCustomAttribute<CommandAttribute>();
                if (attr != null)
                {
                    // 检查返回值类型：必须是 void 或 Task
                    bool isAsyncTask = method.ReturnType == typeof(Task);
                    
                    if (!isAsyncTask && method.ReturnType != typeof(void))
                    {
                        Console.WriteLine($"警告：命令 {attr.Name} 的返回类型必须是 void 或 Task，当前为 {method.ReturnType.Name}");
                        continue;
                    }

                    // 检查是否有 [Profiled] 属性
                    var profiledAttr = method.GetCustomAttribute<ProfiledAttribute>();
                    bool needProfiling = profiledAttr != null;

                    commandInfos[attr.Name] = new CommandInfo
                    {
                        Name = attr.Name,
                        Description = attr.Description,
                        Parameters = attr.Parameters,
                        Group = attr.Group,
                        CommandType = isAsyncTask ? CommandType.Async : CommandType.Sync,  // 标记命令类型
                        AsyncAction = async (string[] args) => 
                        {
                            try
                            {
                                if (isAsyncTask)
                                {
                                    // 异步方法：直接 await
                                    if (needProfiling)
                                    {
                                        var startTime = DateTime.Now;
                                        var task = (Task)method.Invoke(null, [args])!;
                                        await task;
                                        var elapsed = DateTime.Now - startTime;
                                        Console.WriteLine($"\n[性能] 命令 '{attr.Name}' 执行耗时：{elapsed.TotalSeconds:F2}s");
                                    }
                                    else
                                    {
                                        var task = (Task)method.Invoke(null, [args])!;
                                        await task;
                                    }
                                }
                                else
                                {
                                    // 同步方法：直接调用
                                    if (needProfiling)
                                    {
                                        var startTime = DateTime.Now;
                                        method.Invoke(null, [args]);
                                        var elapsed = DateTime.Now - startTime;
                                        Console.WriteLine($"\n[性能] 命令 '{attr.Name}' 执行耗时：{elapsed.TotalSeconds:F2}s");
                                    }
                                    else
                                    {
                                        method.Invoke(null, [args]);
                                    }
                                }
                            }
                            catch (TargetInvocationException ex)
                            {
                                Console.WriteLine($"\n❌ 执行命令 '{attr.Name}' 出错：{ex.InnerException?.Message}");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"\n❌ 调用命令 '{attr.Name}' 失败：{ex.Message}");
                            }
                        }
                    };
                }
            }
            
            asyncCommands = commandInfos.ToDictionary(k => k.Key, v => v.Value.AsyncAction);
        }

        static double CalculateSimilarity(string? query, string? text)
        {
            if (string.IsNullOrEmpty(query) || string.IsNullOrEmpty(text))
            {
                return 0.0;
            }

            // 此时 query 和 text 已被断言为非 null
            string queryLower = query!.ToLower();
            string textLower = text!.ToLower();

            double ratio = CalculateLevenshteinRatio(queryLower, textLower);
            
            if (queryLower.Length > 0 && textLower.Contains(queryLower))
            {
                ratio = Math.Max(ratio, 0.9);
            }

            var querySet = new HashSet<char>(queryLower);
            var textSet = new HashSet<char>(textLower);
            var overlap = querySet.Intersect(textSet).Count() / (double)Math.Max(querySet.Count, 1);

            return 0.6 * ratio + 0.4 * overlap;
        }

        static double CalculateLevenshteinRatio(string s1, string s2)
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

        static void SearchCommands(string keyword, double threshold = 0.3, int? topK = null)
        {
            if (commandInfos == null || commandInfos.Count == 0)
            {
                Console.WriteLine("没有可用的命令。");
                return;
            }

            var results = new List<(string Name, string Group, string Description, double Score)>();
            string keywordLower = (keyword ?? "").ToLower();

            foreach (var cmd in commandInfos.Values)
            {
                string cmdName = cmd.Name ?? "";
                string cmdDesc = cmd.Description ?? "";
                string cmdGroup = cmd.Group ?? "";

                double scoreName = CalculateSimilarity(keyword, cmdName);
                double scoreDesc = CalculateSimilarity(keyword, cmdDesc);
                double scoreGroup = CalculateSimilarity(keyword, cmdGroup);

                double score = Math.Max(scoreName, Math.Max(scoreDesc, scoreGroup));

                if (!string.IsNullOrEmpty(keywordLower) && 
                    (cmdName.ToLower().Contains(keywordLower) || 
                     cmdDesc.ToLower().Contains(keywordLower) || 
                     cmdGroup.ToLower().Contains(keywordLower)))
                {
                    score = Math.Min(1.0, score + 0.5);
                }

                if (score >= threshold)
                {
                    results.Add((cmdName, cmdGroup, cmdDesc, Math.Round(score, 3)));
                }
            }

            results.Sort((a, b) => b.Score.CompareTo(a.Score));

            if (topK.HasValue && topK.Value > 0)
            {
                results = results.Take(topK.Value).ToList();
            }

            Console.WriteLine($"\n找到 {results.Count} 个匹配的命令：\n");
            foreach (var result in results)
            {
                Console.WriteLine($" {result.Name} (相似度：{result.Score})");
                Console.WriteLine($"    分组：{result.Group}");
                if (!string.IsNullOrEmpty(result.Description))
                {
                    Console.WriteLine($"    说明：{result.Description}");
                }
                
                // 显示参数信息
                var cmd = commandInfos.Values.FirstOrDefault(c => c.Name == result.Name);
                if (cmd != null && !string.IsNullOrEmpty(cmd.Parameters))
                {
                    Console.WriteLine($"    参数：{cmd.Parameters}");
                }
            }
        }

    }
}