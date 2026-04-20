using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using cad_tools;
using share.nomal;

namespace tools
{
    class Program
    {
        static SldWorks? swApp;
        static ModelDoc2? swModel;
        static readonly Dictionary<string, Func<string[], Task>> commands = new();

        [STAThread]
        static async Task Main(string[] args)
        {
            swApp = Connect.run();
            if (swApp == null) { Console.WriteLine("无法连接 SolidWorks"); return; }
            swModel = (ModelDoc2)swApp.ActiveDoc;
            SwContext.Instance.Initialize(swApp, swModel);

            RegisterCommands();

            Console.WriteLine("\n输入命令，或描述需求让 AI 处理。输入 help 查看命令，quit 退出。\n");

            while (true)
            {
                Console.Write("> ");
                var input = Console.ReadLine()?.Trim();
                if (string.IsNullOrEmpty(input)) continue;
                if (input == "quit") break;
                if (input == "help") { ShowHelp(); continue; }

                await Execute(input);
            }
        }

        static void RegisterCommands()
        {
            // === 零件命令 ===
            Register("export", "导出零件为 DWG", _ => Exportdwg.run(swModel!, GetThickness()));
            Register("export2", "导出零件为 DWG（版本2）", _ => exportdwg2_body.run(swModel!));
            Register("get_thickness", "获取零件厚度", _ => get_thickness.run(swModel!));
            Register("get_thickness_folder", "从实体文件夹获取厚度", _ => get_thickness_from_solidfolder.run(swModel!));
            Register("add_name", "添加零件名到自定义属性", _ => add_name2info.run(swModel!));
            Register("body_names", "获取所有 body 名称", _ => get_all_body_names.run(swModel!));
            Register("unsuppress", "解除压缩特征", _ => Unsupress.run(swModel!));
            Register("open_folder", "打开零件所在文件夹", _ => open_doc_folder.run(swModel!));
            Register("close_doc", "关闭当前文档", _ => close_current_doc.run(swApp!, swModel!));
            Register("current_name", "获取当前文档名称", _ => Getcurrentdocname.run(swModel!));

            // === 装配体命令 ===
            Register("asm2export", "装配体批量导出 DWG", _ => asm2do.run(swApp!, swModel!, (m, a) => exportdwg2_body.run(m)));
            Register("asm2check", "装配体批量检查展开", _ => asm2do.run(swApp!, swModel!, (m, a) => { checkk_factor.run(a, m); return 0; }));
            Register("asm2bom", "装配体导出 BOM", async _ => await asm2bom.run(swApp!, swModel!,false));
            Register("asm2bomp", "装配体导出零件 BOM", async _ => await asm2bom.run(swApp!, swModel!, true));
            Register("asm2step", "装配体批量导出 STEP", _ => asm2do.run(swApp!, swModel!, (m, a) => one2step.run(m)));
            Register("asm2drw", "装配体批量生成工程图", _ => Asm2Drw());
            Register("all_part_names", "获取所有零件名称", _ => Getallpartname.run(swModel!));
            Register("select_part", "按名称选择零件", args => select_part_byname.run(swModel!, args.Length > 1 ? args[1] : ""));
            Register("open_part", "按名称打开零件", args => { if (args.Length > 1) open_part_by_name.run(swApp!, args[1]); });
            Register("folder2step", "文件夹内零件导出为 STEP", _ => Folder2Step());

            // === 工程图命令 ===
            Register("new_drw", "新建工程图", _ => { add_name2info.run(swModel!); New_drw.run(swApp!, swModel!); });
            Register("drw2dwg", "工程图转 DWG", _ => drw2dwg.run(swModel!, swApp!));
            Register("folderdrw2dwg", "批量转换文件夹工程图为 DWG", _ => FolderDrw2Dwg());
            Register("analyze_face", "分析选中的面", _ => select_face_recognize.run(swModel!));
            Register("drw2png", "工程图转 PNG", _ => drw2png.run(swModel!, swApp!));
            Register("get_edges", "获取工程图所有可见边线", _ => get_all_visable_edge.run(swModel!));
            Register("bend_dim", "自动标注折弯尺寸", _ => benddim.AddBendDimensions(swApp!));

            // === 综合计划 ===
            Register("plan1", "综合计划：添加名称+获取厚度+导出+生成工程图", _ =>
            {
                add_name2info.run(swModel!);
                var t = GetThickness();
                Exportdwg.run(swModel!, t);
                New_drw.run(swApp!, swModel!);
            });

            // === 工作日志命令 ===
            Register("log_add", "添加工作日志 (用法: log_add 工作内容 [备注])", args => AddWorkLog(args));
            Register("log_pending", "查看未完成的工作日志", _ => ShowPendingLogs());
            Register("log_done", "标记工作完成 (用法: log_done ID)", args => SetLogComplete(args, true));
            Register("log_undo", "标记工作未完成 (用法: log_undo ID)", args => SetLogComplete(args, false));
            Register("log_all", "查看所有工作日志", _ => ShowAllLogs());
            Register("log_del", "删除工作日志 (用法: log_del ID)", args => DeleteLog(args));
        }

        static void Register(string name, string desc, Action<string[]> action)
        {
            commands[name.ToLower()] = args => { action(args); return Task.CompletedTask; };
            CommandRegistry.Instance.RegisterCommand(new CommandInfo
            {
                Name = name,
                Description = desc,
                AsyncAction = args => { action(args); return Task.CompletedTask; }
            });
        }

        static void Register(string name, string desc, Func<string[], Task> asyncAction)
        {
            commands[name.ToLower()] = asyncAction;
            CommandRegistry.Instance.RegisterCommand(new CommandInfo
            {
                Name = name,
                Description = desc,
                AsyncAction = asyncAction
            });
        }

        static async Task Execute(string input)
        {
            var parts = input.Split(' ', 2);
            var cmd = parts[0].ToLower();
            var args = parts.Length > 1 ? new[] { parts[0], parts[1] } : new[] { parts[0] };

            if (commands.TryGetValue(cmd, out var action))
            {
                try { await action(args); }
                catch (Exception ex) { Console.WriteLine($"错误: {ex.Message}"); }
            }
            else
            {
                Console.WriteLine($"未知命令: {cmd}");
            }
        }

        static void ShowHelp()
        {
            Console.WriteLine("\n可用命令:");
            foreach (var c in CommandRegistry.Instance.GetAllCommands())
                Console.WriteLine($"  {c.Value.Name} - {c.Value.Description}");
            Console.WriteLine();
        }

        static string GetThickness()
        {
            var t = get_thickness.run(swModel!);
            return t.ToString();
        }


        static void Asm2Drw()
        {
            var names = Getallpartname.run(swModel!);
            if (names == null) return;
            foreach (var name in names)
            {
                open_part_by_name.run(swApp!, name);
                get_thickness.run(swModel!);
                New_drw.run(swApp!, swModel!);
                close_current_doc.run(swApp!, swModel!);
            }
        }

        static void Folder2Step()
        {
            if (swApp == null) return;
            var files = FolderPicker.GetFileNamesFromSelectedFolder();
            if (files == null) return;
            foreach (var file in files)
            {
                if (file.EndsWith(".SLDPRT", StringComparison.OrdinalIgnoreCase))
                    swModel = (ModelDoc2)swApp.OpenDoc6(file, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0);
                else if (file.EndsWith(".SLDASM", StringComparison.OrdinalIgnoreCase))
                    swModel = (ModelDoc2)swApp.OpenDoc6(file, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0);
                else continue;

                if (swModel != null)
                {
                    swModel.Visible = true;
                    one2step.run(swModel);
                    swApp.CloseDoc(swModel.GetTitle());
                    Console.WriteLine($"已转换: {swModel.GetTitle()}");
                }
            }
        }

        static void FolderDrw2Dwg()
        {
            if (swApp == null) return;
            var files = FolderPicker.GetFileNamesFromSelectedFolder();
            if (files == null) return;
            foreach (var file in files)
            {
                if (!file.EndsWith(".SLDDRW", StringComparison.OrdinalIgnoreCase)) continue;
                try
                {
                    var model = (ModelDoc2)swApp.OpenDoc6(file, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0);
                    if (model != null)
                    {
                        drw2dwg.run(model, swApp);
                        swApp.CloseDoc(model.GetTitle());
                        Console.WriteLine($"已转换: {file}");
                    }
                }
                catch (Exception ex) { Console.WriteLine($"转换失败 {file}: {ex.Message}"); }
            }
        }

        // === 工作日志命令实现 ===

        static void AddWorkLog(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("用法: log_add 工作内容 [备注]");
                return;
            }

            // 合并参数作为内容，支持空格
            var input = string.Join(" ", args[1..]);
            var parts = input.Split("//", 2);
            var content = parts[0].Trim();
            var remark = parts.Length > 1 ? parts[1].Trim() : null;

            if (string.IsNullOrWhiteSpace(content))
            {
                Console.WriteLine("工作内容不能为空");
                return;
            }

            var id = WorkLogManager.AddLog(content, remark);
            Console.WriteLine($"已添加工作日志 [ID:{id}]: {content}");
            if (remark != null)
                Console.WriteLine($"备注: {remark}");
        }

        static void ShowPendingLogs()
        {
            var logs = WorkLogManager.GetPendingLogs();
            if (logs.Count == 0)
            {
                Console.WriteLine("暂无未完成的工作日志");
                return;
            }

            Console.WriteLine("\n=== 未完成的工作日志 ===");
            Console.WriteLine($"{"ID",-4} {"创建时间",-20} 工作内容");
            Console.WriteLine(new string('-', 80));
            foreach (var log in logs)
            {
                Console.WriteLine($"{log.Id,-4} {log.CreatedAt:yyyy-MM-dd HH:mm:ss}  {log.Content}");
                if (!string.IsNullOrEmpty(log.Remark))
                    Console.WriteLine($"     备注: {log.Remark}");
            }
            Console.WriteLine($"\n共 {logs.Count} 条未完成日志\n");
        }

        static void ShowAllLogs()
        {
            var logs = WorkLogManager.GetAllLogs();
            if (logs.Count == 0)
            {
                Console.WriteLine("暂无工作日志");
                return;
            }

            Console.WriteLine("\n=== 所有工作日志 ===");
            Console.WriteLine($"{"ID",-4} {"状态",-6} {"创建时间",-20} 工作内容");
            Console.WriteLine(new string('-', 80));
            foreach (var log in logs)
            {
                var status = log.IsCompleted ? "[完成]" : "[待办]";
                Console.WriteLine($"{log.Id,-4} {status,-6} {log.CreatedAt:yyyy-MM-dd HH:mm:ss}  {log.Content}");
                if (!string.IsNullOrEmpty(log.Remark))
                    Console.WriteLine($"     备注: {log.Remark}");
            }
            Console.WriteLine($"\n共 {logs.Count} 条日志\n");
        }

        static void SetLogComplete(string[] args, bool completed)
        {
            if (args.Length < 2 || !int.TryParse(args[1], out var id))
            {
                Console.WriteLine($"用法: {args[0]} ID");
                return;
            }

            var action = completed ? "完成" : "未完成";
            if (WorkLogManager.SetComplete(id, completed))
                Console.WriteLine($"已将工作日志 [ID:{id}] 标记为{action}");
            else
                Console.WriteLine($"未找到工作日志 [ID:{id}]");
        }

        static void DeleteLog(string[] args)
        {
            if (args.Length < 2 || !int.TryParse(args[1], out var id))
            {
                Console.WriteLine("用法: log_del ID");
                return;
            }

            if (WorkLogManager.DeleteLog(id))
                Console.WriteLine($"已删除工作日志 [ID:{id}]");
            else
                Console.WriteLine($"未找到工作日志 [ID:{id}]");
        }
    }
}
