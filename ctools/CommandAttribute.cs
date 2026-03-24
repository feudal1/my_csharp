using System;

namespace tools
{
    [AttributeUsage(AttributeTargets.Method, AllowMultiple = false)]
    public class CommandAttribute : Attribute
    {
        public string Name { get; }
        public string? Description { get; set; }
        public string? Parameters { get; set; }
        public string? Group { get; set; }  // 命令所属组，"cad" 或 "solidworks"
        
        public CommandAttribute(string name)
        {
            Name = name;
        }
    }
}
