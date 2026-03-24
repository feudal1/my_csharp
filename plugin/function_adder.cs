   using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorksTools;
using System;
using System.IO;
using System.Runtime.InteropServices;
using tools;
using System.Diagnostics;
using SolidWorks.Interop.swconst;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Reflection;
using System.Linq;
   namespace SolidWorksAddinStudy
{
   
    public partial class AddinStudy 
{
     // 命令注册表，存储所有通过装饰器注册的命令
     private static Dictionary<int, MethodInfo> _commandRegistry = new Dictionary<int, MethodInfo>();
     
     /// <summary>
     /// 初始化命令注册表，扫描当前类中所有标记了 [Command] 特性的方法
     /// </summary>
     private void InitializeCommandRegistry()
     {
         _commandRegistry.Clear();
         var type = this.GetType();
         
         foreach (var method in type.GetMethods(BindingFlags.NonPublic | BindingFlags.Instance))
         {
             var attr = method.GetCustomAttribute<CommandAttribute>();
             if (attr != null)
             {
                 _commandRegistry[attr.Id] = method;
                 Debug.WriteLine($"[命令注册] ID:{attr.Id}, 名称:{attr.Name}");
             }
         }
     }
     /// <summary>
        /// 通过用户点击的菜单 id 来执行不同的动作
        /// </summary>
        public void FunctionProxy(string data)
        {
            int commandId = int.Parse(data);

            if (_commandRegistry.TryGetValue(commandId, out var method))
            {
                try
                {
                    var cmdAttr = method.GetCustomAttribute<CommandAttribute>();
                    
                    // 如果设置了 ShowOutputWindow，则包装执行
                    if (cmdAttr != null && cmdAttr.ShowOutputWindow)
                    {
                        ShowOutputWindow();
                        method.Invoke(this, null);
                    }
                    else
                    {
                        method.Invoke(this, null);
                    }
                }
                catch (Exception ex)
                {
                    swApp?.SendMsgToUser($"执行命令失败：{ex.Message}");
                }
            }
            else
            {
                swApp?.SendMsgToUser($"未找到命令 ID: {commandId}");
            }
        }
   private void AddCommandMgr()
        {
            try
            {
                InitializeCommandRegistry();
                
                int mainCmdGroupID = 5001;
                int[] mainItemIds = new List<int>(_commandRegistry.Keys).ToArray();

                int cmdGroupErr = 0;
                bool ignorePrevious = false;

                object registryIDs = null;
                bool getDataResult = iCmdMgr?.GetGroupDataFromRegistry(mainCmdGroupID, out registryIDs) ?? false;

                if (getDataResult && registryIDs != null && registryIDs is int[] regIds)
                {
                    if (!CompareIDs(regIds, mainItemIds))
                    {
                        ignorePrevious = true;
                    }
                }

                // 强制忽略之前的配置，确保重新创建所有 CommandTab
                ignorePrevious = true;

                var cmdGroup = iCmdMgr?.CreateCommandGroup2(
                    mainCmdGroupID,
                    "调试工具",
                    "调试控制台控制",
                    "",
                    -1,
                    ignorePrevious,
                    ref cmdGroupErr
                );

                if (cmdGroup == null) return;

                string[] icons = new string[6];
                for (int i = 0; i < 6; i++)
                {
                    icons[i] = "";
                }
                cmdGroup.IconList = icons;

                int menuToolbarOption = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);

                foreach (var kvp in _commandRegistry)
                {
                    var cmdAttr = kvp.Value.GetCustomAttribute<CommandAttribute>();
                    if (cmdAttr != null)
                    {
                        int cmdIndex = cmdGroup.AddCommandItem2(
                            cmdAttr.Name,
                            -1,
                            cmdAttr.Tooltip,
                            cmdAttr.LocalizedName,
                            0,
                            $"FunctionProxy({cmdAttr.Id})",
                            $"EnableFunction({cmdAttr.Id})",
                            kvp.Key,
                            menuToolbarOption
                        );
                    }
                }

                cmdGroup.HasToolbar = true;
                cmdGroup.HasMenu = true;
                cmdGroup.Activate();

                // 按文档类型分组命令
                var allDocTypes = new[] { 
                    (int)swDocumentTypes_e.swDocPART, 
                    (int)swDocumentTypes_e.swDocASSEMBLY, 
                    (int)swDocumentTypes_e.swDocDRAWING 
                };

                foreach (var docType in allDocTypes)
                {
                    // 获取该文档类型应该显示的命令
                    var commandsForDocType = _commandRegistry.Values
                        .Select(m => m.GetCustomAttribute<CommandAttribute>())
                        .Where(attr => attr != null && 
                              (attr.DocumentTypes.Contains(0) || attr.DocumentTypes.Contains(docType)))
                        .ToArray();

                    if (commandsForDocType.Length == 0) continue;

                    string docTypeName = docType == (int)swDocumentTypes_e.swDocPART ? "零件" :
                                        docType == (int)swDocumentTypes_e.swDocASSEMBLY ? "装配体" : "工程图";
                    
                    Debug.WriteLine($"[文档类型:{docTypeName}] 命令数：{commandsForDocType.Length}");
                    foreach (var cmd in commandsForDocType)
                    {
                        Debug.WriteLine($"  - {cmd.Name} (ID:{cmd.Id}, 文档类型:[{string.Join(",", cmd.DocumentTypes)}])");
                    }

                    var cmdTab = iCmdMgr?.GetCommandTab(docType, "调试工具");

                    if (cmdTab != null && (!getDataResult || ignorePrevious))
                    {
                        iCmdMgr?.RemoveCommandTab(cmdTab);
                        cmdTab = null;
                    }

                    if (cmdTab == null)
                    {
                        cmdTab = iCmdMgr?.AddCommandTab(docType, "调试工具");

                        var cmdBox = cmdTab?.AddCommandTabBox();

                        List<int> cmdIDs = new List<int>();
                        List<int> showTextType = new List<int>();

                        foreach (var cmdAttr in commandsForDocType)
                        {
                            int cmdIndexInGroup = Array.IndexOf(mainItemIds, cmdAttr.Id);
                            cmdIDs.Add(cmdGroup.get_CommandID(cmdIndexInGroup));
                            showTextType.Add((int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow);
                        }

                        cmdBox?.AddCommands(cmdIDs.ToArray(), showTextType.ToArray());
                    }
                }
            }
            catch (Exception ex)
            {
                swApp?.SendMsgToUser($"添加按钮失败：{ex.Message}");
            }
        }

}}