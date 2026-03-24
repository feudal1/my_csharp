using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Text;
using SolidWorks.Interop.sldworks;

namespace tools
{
    /// <summary>
    /// 命令执行器 - 用于解析和执行命令
    /// </summary>
    public class CommandExecutor
    {
        private readonly Func<string, CommandInfo?> _commandResolver;
        private readonly Func<SldWorks?> _swAppResolver;
        
        public CommandExecutor(
            Func<string, CommandInfo?> commandResolver,
            Func<SldWorks?> swAppResolver)
        {
            _commandResolver = commandResolver;
            _swAppResolver = swAppResolver;
        }
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
                    args = parts.Length > 1 ? parts.Skip(1).ToArray() : new string[0];
                }

                // 检查命令是否存在
                var commandInfo = _commandResolver(commandName);
                if (commandInfo == null)
                {
                    return $"错误：未找到命令 '{commandName}'。请使用 search 命令查看可用命令。";
                }

                // 检查是否连接到 SolidWorks
                var swApp = _swAppResolver();
                if (swApp == null)
                {
                    return "错误：未连接到 SolidWorks，请先启动程序";
                }

                // 执行命令
                await commandInfo!.AsyncAction(args);

                return $"命令 '{commandName}' 执行成功";
            }
            catch (Exception ex)
            {
                return $"执行命令失败：{ex.Message}";
            }
        }
    }
}
