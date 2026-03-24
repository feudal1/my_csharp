using System.Threading.Tasks;

namespace tools
{
    /// <summary>
    /// 命令类型枚举
    /// </summary>
    public enum CommandType
    {
        Sync,      // 同步命令
        Async      // 异步命令
    }

    /// <summary>
    /// 命令信息类
    /// </summary>
    public class CommandInfo
    {
        public string Name { get; set; } = "";
        public string? Description { get; set; }
        public string? Parameters { get; set; }
        public string? Group { get; set; }
        public Func<string[], Task> AsyncAction { get; set; } = null!;
        public CommandType CommandType { get; set; } = CommandType.Sync;
        
        /// <summary>
        /// 执行命令（支持参数）
        /// </summary>
        public async Task ExecuteAsync(string[] args)
        {
            if (AsyncAction == null)
            {
                throw new InvalidOperationException($"命令 {Name} 未绑定执行方法");
            }
            
            await AsyncAction(args);
        }
    }
}
