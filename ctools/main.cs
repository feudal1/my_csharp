using System;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

using System.Collections.Generic;
using System.Reflection;
using System.Linq;
using System.Text.Json;
using System.Diagnostics;
using System.Text;
using tools;

namespace tools
{
    // Windows API 用于设置控制台代码页
    class ConsoleHelper
    {
        [DllImport("kernel32.dll")]
        public static extern bool SetConsoleCP(uint wCodePageID);

        [DllImport("kernel32.dll")]
        public static extern bool SetConsoleOutputCP(uint wCodePageID);
    }
    
    // 性能监控装饰器属性
    [System.AttributeUsage(System.AttributeTargets.Method)]
    class ProfiledAttribute : System.Attribute
    {
        public string? Description { get; set; }
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
        
        static Dictionary<string, CommandInfo>? commandInfos;
               static private DateTime GetBuildTime()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            string filePath = assembly.Location;
            return new FileInfo(filePath).LastWriteTime;
        }
        
        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine($"ctools build time: {GetBuildTime():MM-dd HH:mm}");
            RegisterCommands();
            
                // 将 ctool 的命令注册到全局命令注册中心
                CommandRegistry.Instance.RegisterAssembly(typeof(Program).Assembly);
                
            if (args.Length==0)
            {
                // 连接 SolidWorks
                swApp = Connect.run();
                if (swApp == null)
                {
                    Console.WriteLine("错误：无法连接到 SolidWorks 应用程序。");
                    return;
                }
                
                swModel = (ModelDoc2)swApp.ActiveDoc;
                
                // 初始化全局 SolidWorks 上下文
                SwContext.Instance.Initialize(swApp, swModel);
                
                // 创建 LLM 循环调用器，注入命令解析器和 SolidWorks 实例解析器
                LlmLoopCaller loopCaller = new LlmLoopCaller(
                    // 传入获取命令描述内容的委托（实时生成）
                    () => GetCommandsDescriptionContent(),
                    // 命令解析器：从全局注册中心查找命令
                    commandName => CommandRegistry.Instance.GetCommand(commandName),
                    // SolidWorks 实例解析器
                    () => swApp,
                    // swModel 更新器
                    (model) => swModel = model
                );
           
                var task = Task.Run(() => loopCaller.InteractiveLoopAsync());
                task.GetAwaiter().GetResult();
            }
            else
            {
                // 如果有命令行参数，显示帮助信息
                Console.WriteLine("\n欢迎使用 ctools - SolidWorks 智能助手");
                Console.WriteLine("\n使用方法:");
                Console.WriteLine("  ctool.exe                  - 启动交互式对话模式");
                Console.WriteLine("\n在对话中你可以:");
                Console.WriteLine("  - 直接输入命令名称执行 (如 exportdxf)");
                Console.WriteLine("  - 用自然语言描述需求 (AI 会自动识别并调用命令)");
                Console.WriteLine("  - 输入 clear 清空历史、history 查看历史、mode 切换模式");
                Console.WriteLine("  - 输入 llm 进入纯对话模式、quit/exit 退出程序");
                Console.WriteLine("\n示例:");
                Console.WriteLine("  > exportdxf              # 直接执行导出 DXF 命令");
                Console.WriteLine("  > 导出当前零件           # AI 会识别意图并调用相应命令");
                Console.WriteLine("\n");
                ShowHelp();
            }
        }
        
        /// <summary>
        /// 从全局命令注册中心获取命令描述内容（实时生成）
        /// </summary>
        static string GetCommandsDescriptionContent()
        {
            var sb = new StringBuilder();
            sb.AppendLine("=== SolidWorks 自动化命令列表 ===");
            sb.AppendLine($"更新时间：{DateTime.Now:yyyy-MM-dd HH:mm:ss}");

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
                    }
                    else
                    {
                        sb.AppendLine($"    参数：{cmd.Parameters}");
                    }
                }
            }
            
            return sb.ToString();
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