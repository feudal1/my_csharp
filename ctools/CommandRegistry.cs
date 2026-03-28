using System;
using System.Collections.Generic;
using System.Reflection;
using System.Threading.Tasks;

namespace tools
{
    /// <summary>
    /// 全局命令注册中心（单例模式）
    /// 用于在不同应用程序间共享命令信息
    /// </summary>
    public class CommandRegistry
    {
        private static readonly Lazy<CommandRegistry> _instance = 
            new Lazy<CommandRegistry>(() => new CommandRegistry());
        
        public static CommandRegistry Instance => _instance.Value;
        
        private Dictionary<string, CommandInfo> _commands = new Dictionary<string, CommandInfo>();
        private object _lock = new object();
        
        /// <summary>
        /// 私有构造函数（单例模式）
        /// </summary>
        private CommandRegistry()
        {
        }
        
        /// <summary>
        /// 注册单个命令
        /// </summary>
        public void RegisterCommand(CommandInfo commandInfo)
        {
            lock (_lock)
            {
                if (string.IsNullOrEmpty(commandInfo.Name))
                {
                    throw new ArgumentException("命令名称不能为空");
                }
                
                _commands[commandInfo.Name.ToLower()] = commandInfo;
                
                // 注册别名
                if (commandInfo.Aliases != null)
                {
                    foreach (var alias in commandInfo.Aliases)
                    {
                        if (!string.IsNullOrEmpty(alias))
                        {
                            _commands[alias.ToLower()] = commandInfo;
                            Console.WriteLine($"[调试] 注册命令别名：{alias} -> {commandInfo.Name}");
                        }
                    }
                }
            }
        }
        
        /// <summary>
        /// 从程序集批量注册命令（通过反射扫描 [Command] 特性）
        /// </summary>
        public void RegisterAssembly(Assembly assembly)
        {
            try
            {
                // 获取所有静态方法（包括公共和非公共）
                var methods = assembly.GetTypes()
                    .SelectMany(t => t.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static));
                
                foreach (var method in methods)
                {
                    var attr = method.GetCustomAttribute<CommandAttribute>();
                    if (attr != null)
                    {
                        var commandInfo = CreateCommandInfoFromAttribute(method, attr);
                        RegisterCommand(commandInfo);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"从程序集 {assembly.GetName().Name} 注册命令失败：{ex.Message}");
            }
        }
        
        /// <summary>
        /// 从类型实例注册命令（用于插件等实例方法）
        /// </summary>
        public void RegisterType(object instance, Type type)
        {
            try
            {
                var methods = type.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                
                foreach (var method in methods)
                {
                    var attr = method.GetCustomAttribute<CommandAttribute>();
                    if (attr != null)
                    {
                        var commandInfo = CreateCommandInfoFromInstanceMethod(instance, method, attr);
                        RegisterCommand(commandInfo);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"从类型 {type.Name} 注册命令失败：{ex.Message}");
            }
        }
        
        /// <summary>
        /// 根据命令名查找 CommandInfo
        /// </summary>
        public CommandInfo? GetCommand(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                Console.WriteLine($"[调试] CommandRegistry.GetCommand - 名称为空");
                return null;
            }
            
            lock (_lock)
            {
                bool found = _commands.TryGetValue(name.ToLower(), out var commandInfo);
                Console.WriteLine($"[调试] CommandRegistry.GetCommand('{name}') - 找到：{found}");
                if (found && commandInfo != null)
                {
                    Console.WriteLine($"[调试] 命令信息 - Name: {commandInfo.Name}, Type: {commandInfo.CommandType}, Group: {commandInfo.Group}");
                }
                return commandInfo;
            }
        }
        
        /// <summary>
        /// 获取所有已注册的命令
        /// </summary>
        public Dictionary<string, CommandInfo> GetAllCommands()
        {
            lock (_lock)
            {
                return new Dictionary<string, CommandInfo>(_commands);
            }
        }
        
        /// <summary>
        /// 清空所有注册的命令
        /// </summary>
        public void Clear()
        {
            lock (_lock)
            {
                _commands.Clear();
            }
        }
        
        /// <summary>
        /// 从 CommandAttribute 创建 CommandInfo（静态方法）
        /// </summary>
        private CommandInfo CreateCommandInfoFromAttribute(MethodInfo method, CommandAttribute attr)
        {
            bool isAsyncTask = method.ReturnType == typeof(Task);
            
            return new CommandInfo
            {
                Name = attr.Name,
                Description = attr.Description,
                Parameters = attr.Parameters,
                Group = attr.Group,
                Aliases = attr.Aliases,
                CommandType = isAsyncTask ? CommandType.Async : CommandType.Sync,
                AsyncAction = async (args) =>
                {
                    try
                    {
                        if (isAsyncTask)
                        {
                            var task = (Task)method.Invoke(null, new object[] { args })!;
                            await task;
                        }
                        else
                        {
                            method.Invoke(null, new object[] { args });
                        }
                    }
                    catch (TargetInvocationException ex)
                    {
                        Console.WriteLine($"\n❌ 执行命令 '{attr.Name}' 出错：{ex.InnerException?.Message}");
                        throw;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"\n❌ 调用命令 '{attr.Name}' 失败：{ex.Message}");
                        throw;
                    }
                }
            };
        }
        
        /// <summary>
        /// 从实例方法创建 CommandInfo（用于插件等需要实例的方法）
        /// </summary>
        private CommandInfo CreateCommandInfoFromInstanceMethod(object instance, MethodInfo method, CommandAttribute attr)
        {
            bool isAsyncTask = method.ReturnType == typeof(Task);
            
            return new CommandInfo
            {
                Name = attr.Name,
                Description = attr.Description,
                Parameters = attr.Parameters,
                Group = attr.Group,
                Aliases = attr.Aliases,
                CommandType = isAsyncTask ? CommandType.Async : CommandType.Sync,
                AsyncAction = async (args) =>
                {
                    try
                    {
                        if (isAsyncTask)
                        {
                            var task = (Task)method.Invoke(instance, null)!;
                            await task;
                        }
                        else
                        {
                            method.Invoke(instance, null);
                        }
                    }
                    catch (TargetInvocationException ex)
                    {
                        Console.WriteLine($"\n❌ 执行命令 '{attr.Name}' 出错：{ex.InnerException?.Message}");
                        throw;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"\n❌ 调用命令 '{attr.Name}' 失败：{ex.Message}");
                        throw;
                    }
                }
            };
        }
    }
}
