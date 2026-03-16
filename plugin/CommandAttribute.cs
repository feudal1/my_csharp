using System;

namespace SolidWorksAddinStudy
{
    /// <summary>
    /// 命令特性标记，用于声明式注册 SolidWorks 插件命令
    /// </summary>
    [AttributeUsage(AttributeTargets.Method, AllowMultiple = false)]
    public class CommandAttribute : Attribute
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Tooltip { get; set; }
        public string LocalizedName { get; set; }
        public int[] DocumentTypes { get; set; }
        
        public CommandAttribute(int id, string name, string tooltip, string localized_name, params int[] documentTypes)
        {
            Id = id;
            Name = name;
            Tooltip = tooltip;
            LocalizedName = localized_name;
            DocumentTypes = documentTypes?.Length > 0 ? documentTypes : new[] { 0 }; 
        }
    }
}