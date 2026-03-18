using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Text;

namespace tools
{
    /// <summary>
    /// 命令执行器 - 用于解析和执行 main.cs 中注册的命令
    /// </summary>
    public class CommandExecutor
    {
        /// <summary>
        /// 解析并执行命令
        /// </summary>
        /// <param name="commandText">完整命令文本，格式："do_【命令名】参数 1 参数 2" 或 "命令名 参数 1 参数 2"</param>
        /// <returns>执行结果</returns>
        public async Task<string> ExecuteCommandAsync(string commandText)
        {
            if (string.IsNullOrWhiteSpace(commandText))
            {
                return "错误：命令不能为空";
            }

            try
            {
                string commandName = commandText;
                string[] args = new string[0];

                // 检查是否包含 do_【】格式
                var strictPattern = new System.Text.RegularExpressions.Regex(@"do_【([^】]+)】\s*(.*)");
                var strictMatch = strictPattern.Match(commandText);
                
                if (strictMatch.Success)
                {
                    // 严格格式：do_【命令名】参数
                    commandName = strictMatch.Groups[1].Value.Trim();
                    string parameters = strictMatch.Groups[2].Value.Trim();
                    
                    // 只有当参数字符串非空时才解析参数
                    if (!string.IsNullOrEmpty(parameters))
                    {
                        args = parameters.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                    }
                    else
                    {
                        args = new string[0];
                    }
                }
                else
                {
                    // 普通格式：命令名 参数 1 参数 2
                    var parts = commandText.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length == 0)
                    {
                        return "错误：无法解析命令";
                    }
                    commandName = parts[0];
                    args = parts.Length > 1 ? parts[1..] : new string[0];
                }

                // 检查命令是否存在
                if (Program.Commands == null || !Program.Commands.ContainsKey(commandName))
                {
                    return $"错误：未找到命令 '{commandName}'。请使用 search 命令查看可用命令。";
                }

                // 检查是否连接到 SolidWorks
                if (Program.SwApp == null)
                {
                    return "错误：未连接到 SolidWorks，请先启动程序";
                }

                // 执行命令
                var commandInfo = Program.Commands[commandName];
                await commandInfo.AsyncAction(args);

                return $"命令 '{commandName}' 执行成功";
            }
            catch (Exception ex)
            {
                return $"执行命令失败：{ex.Message}";
            }
        }

        /// <summary>
        /// 获取所有可用命令的描述
        /// </summary>
        /// <returns>命令描述文本</returns>
        public static string GetAllCommandsDescription()
        {
            if (Program.Commands == null || Program.Commands.Count == 0)
            {
                return "暂无可用命令";
            }

            var sb = new StringBuilder();
            foreach (var cmd in Program.Commands.Values)
            {
                sb.AppendLine($"\n【{cmd.Group ?? "默认"}】{cmd.Name} {(cmd.CommandType == CommandType.Async ? "(异步)" : "")}");
                if (!string.IsNullOrEmpty(cmd.Description))
                {
                    sb.AppendLine($"    说明：{cmd.Description}");
                }
                if (!string.IsNullOrEmpty(cmd.Parameters))
                {
                    sb.AppendLine($"    参数：{cmd.Parameters}");
                }
            }

            return sb.ToString();
        }
    }
}
