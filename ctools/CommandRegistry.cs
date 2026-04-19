using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace tools
{
    public class CommandRegistry
    {
        private static readonly CommandRegistry _instance = new();
        public static CommandRegistry Instance => _instance;
        private readonly Dictionary<string, CommandInfo> _commands = new();

        public void RegisterCommand(CommandInfo cmd)
        {
            _commands[cmd.Name.ToLower()] = cmd;
        }

        public CommandInfo? GetCommand(string name)
        {
            _commands.TryGetValue(name.ToLower(), out var cmd);
            return cmd;
        }

        public Dictionary<string, CommandInfo> GetAllCommands() => new(_commands);
    }

    public class CommandInfo
    {
        public string Name { get; set; } = "";
        public string? Description { get; set; }
        public Func<string[], Task> AsyncAction { get; set; } = _ => Task.CompletedTask;
    }
}
