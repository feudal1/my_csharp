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
        private readonly Action<ModelDoc2?> _swModelUpdater;
        
        public CommandExecutor(
            Func<string, CommandInfo?> commandResolver,
            Func<SldWorks?> swAppResolver,
            Action<ModelDoc2?> swModelUpdater)
        {
            _commandResolver = commandResolver;
            _swAppResolver = swAppResolver;
            _swModelUpdater = swModelUpdater;
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

                // 普通格式：命令名 参数 1 参数 2
                var parts = commandText.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 0)
                {
                    return "错误：无法解析命令";
                }
                commandName = parts[0];
                args = parts.Length > 1 ? parts.Skip(1).ToArray() : new string[0];

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
                    Console.WriteLine("\n[错误] CommandExecutor: SolidWorks 未连接");
                    return "错误：未连接到 SolidWorks，请先启动程序";
                }

                // 每次执行命令前重新获取当前激活的模型
                var swModel = (ModelDoc2)swApp.ActiveDoc;
                
                // 如果 ActiveDoc 为 null，尝试通过 IActiveDoc2 获取
                if (swModel == null)
                {
                    try
                    {
                        swModel = (ModelDoc2)swApp.IActiveDoc2;
                        Console.WriteLine($"\n[调试] CommandExecutor: 通过 IActiveDoc2 获取到模型：{(swModel != null ? swModel.GetTitle() : "null")}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"\n[调试] CommandExecutor: IActiveDoc2 获取失败：{ex.Message}");
                    }
                }
                
                _swModelUpdater(swModel);
                
                if (swModel == null)
                {
                    Console.WriteLine("\n[警告] CommandExecutor: 当前没有激活的文档");
                }
                else
                {
                    Console.WriteLine($"\n[调试] CommandExecutor: 当前模型：{swModel.GetTitle()}");
                }

                // 执行命令
                Console.WriteLine($"\n[调试] CommandExecutor: 即将执行命令：{commandName}");
                Console.WriteLine($"[调试] CommandExecutor: 参数数量：{args.Length}");
                Console.WriteLine($"[调试] CommandExecutor: commandInfo.CommandType = {commandInfo.CommandType}");
                
                await commandInfo!.AsyncAction(args);

                Console.WriteLine($"[调试] CommandExecutor: 命令 {commandName} 执行完成");

                return $"命令 '{commandName}' 执行成功";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n[错误] CommandExecutor: 命令执行异常：{ex}");
                Console.WriteLine($"[错误] StackTrace: {ex.StackTrace}");
                return $"执行命令失败：{ex.Message}";
            }
        }
    }
}
