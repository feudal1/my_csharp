using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using tools;
using cad_tools;

namespace tools
{
    partial class Program
    {
        [Command("export", Description = "导出当前零件展开为 DWG 文件", Parameters = "无", Group = "solidworks")]
        static void ExportCommand(string[] args)
        {
            if (swApp == null || swModel == null) return;

            var thickness = get_thickness.run(swModel);
            var dwgname = Exportdwg.run(swModel, thickness.ToString());
      
        }
        [Command("get_select_type", Description = "导出当前零件展开为 DWG 文件", Parameters = "无", Group = "solidworks")]
        static void Get_select_type(string[] args)
        {
            if (swApp == null || swModel == null) return;

            get_select_type.run(SwModel);

        }
        [Command("export_flat_view", Description = "导出钣金展开视图 dwg", Parameters = "无", Group = "solidworks")]
        static void ExportFlatViewCommand(string[] args)
        {
            if (swApp == null || swModel == null) return;
        
            var thickness = get_thickness.run(swModel);
            Exportdwg.run(swModel, thickness.ToString());
        }


        [Command("asm2export", Description = "装配体批量导出 dwg", Parameters = "无", Group = "solidworks")]
        static void Asm2Export2Command(string[] args)
        {
            if (swApp == null || swModel == null) return;

            asm2do.run(swApp, swModel, (model, app) =>
            {
                checkk_factor.run(app, model);
                return exportdwg2_body.run(model);
            });
        }
        [Command("asm2check", Description = "装配体批量检查展开", Parameters = "无", Group = "solidworks")]
        static void Asm2check(string[] args)
        {
            if (swApp == null || swModel == null) return;

            asm2do.run(swApp, swModel, (model, app) => {
                checkk_factor.run(app,model);
                return 0;
            });
        }
        [Command("get_all_typename", Description = "获取所有零件名称", Parameters = "无", Group = "solidworks")]
        static void GetAllTypename(string[] args)
        {
            if (swApp == null || swModel == null) return;

            get_all_typename.run( swModel);
        }
        [Command("asm2bom", Description = "装配体导出bom", Parameters = "无", Group = "solidworks")]
        static void Asm2bom(string[] args)
        {
            if (swApp == null || swModel == null) return;

            asm2do.run(swApp, swModel, (model, app) => {
                asm2bom.run(app,model);
                return 0;
            });
        }

        [Command("asm2drw", Description = "装配体批量生成工程图", Parameters = "无", Group = "solidworks")]
        [Profiled(Description = "性能监控：装配体批量生成工程图")]
        static void Asm2DrwCommand(string[] args)
        {
            if (swApp == null || swModel == null) return;

            string[]? partnames = Getallpartname.run(swModel);
            if (partnames != null && partnames.Length > 0)
            {
                foreach (var partname in partnames)
                {
                    Profiler.Time(() => open_part_by_name.run(swApp, partname));
                    var thickness = Profiler.Time(() => get_thickness.run(swModel));
                    Profiler.Time(() => New_drw.run(swApp, swModel));
                    Profiler.Time(() => close_current_doc.run(swApp, swModel));
                }
            }
        }

        [Command("unsuppress", Description = "解除压缩特征或零部件", Parameters = "无", Group = "solidworks")]
        static void UnsuppressCommand(string[] args)
        {
            if (swModel == null) return;

            Unsupress.run(swModel);
        }

        [Command("get_select_part_name", Description = "获取选中零件的名称", Parameters = "需先选择零件", Group = "solidworks")]
        static void GetSelectPartNameCommand(string[] args)
        {
            if (swModel == null) return;

            Getselectpartname.run(swModel);
        }

        [Command("get_all_part_name", Description = "获取装配体中所有零件名称", Parameters = "无", Group = "solidworks")]
        static void GetAllPartNameCommand(string[] args)
        {
            if (swModel == null) return;

            Getallpartname.run(swModel);
        }

        [Command("open_part_by_name", Description = "按名称打开零件", Parameters = "<零件名称>", Group = "solidworks")]
        static void OpenPartByNameCommand(string[] args)
        {
            if (swApp == null) return;

            if (args.Length > 1)
            {
                open_part_by_name.run(swApp, args[1]);
            }
            else
            {
                Console.WriteLine("请提供零件名称参数");
            }
        }

        [Command("new_drw", Description = "新建工程图", Parameters = "无", Group = "solidworks")]
        static void NewDrwCommand(string[] args)
        {
            if (swApp == null || swModel == null) return;
            add_name2info.run(swModel);
            New_drw.run(swApp, swModel);
        }

        [Command("get_current_doc_name", Description = "获取当前文档名称", Parameters = "无", Group = "solidworks")]
        static void GetCurrentDocNameCommand(string[] args)
        {
            if (swModel == null) return;

            Getcurrentdocname.run(swModel);
        }

        [Command("close_current_doc", Description = "关闭当前文档", Parameters = "无", Group = "solidworks")]
        static void CloseCurrentDocCommand(string[] args)
        {
            if (swApp == null || swModel == null) return;

            close_current_doc.run(swApp, swModel);
        }

        [Command("get_thickness", Description = "获取零件厚度", Parameters = "无", Group = "solidworks")]
        static void GetThicknessCommand(string[] args)
        {
            if (swModel == null) return;

            get_thickness.run(swModel);
        }

        [Command("get_thickness_from_solidfolder", Description = "从实体文件夹获取厚度", Parameters = "无", Group = "solidworks")]
        static void GetThicknessFromSolidFolderCommand(string[] args)
        {
            if (swModel == null) return;

            get_thickness_from_solidfolder.run(swModel);
        }

        [Command("open_doc_folder", Description = "打开零件所在文件夹", Parameters = "无", Group = "solidworks")]
        static void OpenDocFolderCommand(string[] args)
        {
            if (swModel == null) return;

            open_doc_folder.run(swModel);
        }

        [Command("plan1", Description = "综合计划：添加名称 + 获取厚度 + 导出 dwg + 生成工程图", Parameters = "无", Group = "solidworks")]
        static void Plan1Command(string[] args)
        {
            if (swApp == null || swModel == null) return;

            add_name2info.run(swModel);
            var thickness = get_thickness.run(swModel);
            NativeClipboard.SetText(thickness.ToString() + "厘");
            var dwgname = Exportdwg.run(swModel, thickness.ToString());
            New_drw.run(swApp, swModel);
        }



        [Command("export2", Description = "导出当前零件展开为 DWG 文件（版本 2）", Parameters = "无", Group = "solidworks")]
        static void Export2Command(string[] args)
        {
            if (swApp == null || swModel == null) return;

             var thickness = get_thickness.run(swModel);
            exportdwg2_body.run(swModel);
        }
             [Command("drw2dwg", Description = "将工程图转换为 DWG 格式", Parameters = "无", Group = "solidworks")]
        static void Drw2DwgCommand(string[] args)
        {
            if (swApp == null || swModel == null) return;

            var dwgFileName =drw2dwg.run(swModel, swApp);
            open_cad_doc_by_name.run(dwgFileName);
        }
               [Command("folderdrw2dwg", Description = "批量转换文件夹内工程图为 DWG 格式", Parameters ="无", Group = "solidworks")]
        static void FolderDrw2DwgCommand(string[] args)
        {
            if (swApp == null) return;

            var files = FolderPicker.GetFileNamesFromSelectedFolder();
            if (files != null)
            {
                foreach (var file in files)
                {
                        // 筛选 SLDDRW 后缀的文件
                    if (!file.EndsWith(".SLDDRW", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    try
                    {
                        // 打开工程图文件
                        ModelDoc2 swModel = (ModelDoc2)swApp.OpenDoc6(
                            file, 
                            (int)swDocumentTypes_e.swDocDRAWING, 
                            (int)swOpenDocOptions_e.swOpenDocOptions_Silent, 
                            "", 
                            0, 
                            0);

                        if (swModel != null)
                        {
                            // 转换为 DWG
                            var dwgFileName = drw2dwg.run(swModel, swApp);
                            
                            // 关闭已处理的文档
                            swApp.CloseDoc(swModel.GetTitle());
                                
            
            
                            Console.WriteLine($"已转换：{dwgFileName}");
                        }
                        else
                        {
                            Console.WriteLine($"无法打开文件：{file}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"转换失败 {file}: {ex.Message}");
                    }
                }
            }
            
            
        }

                [Command("add_name2info", Description = "添加零件名称到自定义属性", Parameters = "无", Group = "solidworks")]
        static void AddName2InfoCommand(string[] args)
        {
            if (swModel == null) return;

            add_name2info.run(swModel);
        }
           
        [Command("drw2png", Description = "将工程图转换为 NPG 格式", Parameters = "无", Group = "solidworks")]
        static void Drw2NpgCommand(string[] args)
        {
            if (swApp == null || swModel == null) return;

            var npgFileName = drw2png.run(swModel, swApp);
        }
        [Command("drw2png2vlm", Description = "导出工程图为 PNG 并调用 VLM 分析", Parameters = "无", Group = "solidworks")]
        static async Task Drw2PngVlmCommand(string[] args)
    {
        // 1. 基础检查
        if (swApp == null || swModel == null)
        {
            Console.WriteLine("错误：未检测到 SolidWorks 应用程序或当前模型。请先打开一个图纸。");
            return;
        }

        Console.WriteLine(">>> 开始导出 PNG...");
        // 假设 drw2png 是你的另一个类
        string pngPath = drw2png.run(swModel, swApp);

        if (string.IsNullOrEmpty(pngPath) || !File.Exists(pngPath))
        {
            Console.WriteLine("错误：PNG 导出失败，文件未生成。");
            return;
        }
        Console.WriteLine($"成功导出: {pngPath}");

        // 2. 获取 API Key (建议在读取 Prompt 之前，避免用户输入了 Prompt 才发现没 Key)
       

        // 3. 获取用户提示语
        Console.WriteLine("\n=== 请输入分析提示语 ===");
        Console.Write("Prompt (直接回车使用默认): ");
        string? promptInput = Console.ReadLine();
        string prompt = string.IsNullOrWhiteSpace(promptInput) ? "请详细分析这张工程图，提取关键尺寸、公差、材料及技术要求。" : promptInput;

        // 4. 调用 VLM (流式)
        Console.WriteLine("\n正在连接 AI 服务，实时生成回复...\n----------------------------");
        
        try
        {
            // 确保 VlmService 是静态类或者你已经实例化了它
            // 如果 VlmService 是静态类且方法也是静态的：
            LlmService vlmService = new LlmService();
            await vlmService.ChatAsync( prompt,pngPath);
            
            Console.WriteLine("\n----------------------------");
            Console.WriteLine("✅ 分析完成。");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"\n❌ VLM 调用失败: {ex.Message}");
            // 可选：记录详细日志
            // Console.WriteLine(ex.StackTrace);
        }
    }

 [Command("get_all_visable_edge", Description = "获取工程图中所有可见边线", Parameters = "无", Group = "solidworks")]
        static void GetAllVisableEdgeCommand(string[] args)
        {
            if (swModel == null) return;

            get_all_visable_edge.run(swModel);
        }

        [Command("asm2step", Description = "装配体导出为 STEP 格式", Parameters = "无", Group = "solidworks")]
        static void Asm2StepCommand(string[] args)
        {
            if (swApp == null || swModel == null) return;

            asm2step.run(swModel);
        }
        [Command("folder2step", Description = "文件夹内零件装配体导出为 STEP 格式", Parameters = "无", Group = "solidworks")]
        static void folder2StepCommand(string[] args)
        {
            if (swApp == null) return;

            var files = FolderPicker.GetFileNamesFromSelectedFolder();
            if (files != null)
            {
                foreach (var file in files)
                {
                    // 筛选 SLDDRW 后缀的文件
                    if (file.EndsWith(".SLDPRT", StringComparison.OrdinalIgnoreCase))
                    {
                        // 打开工程图文件
                         swModel = (ModelDoc2)swApp.OpenDoc6(
                            file, 
                            (int)swDocumentTypes_e.swDocPART, 
                            (int)swOpenDocOptions_e.swOpenDocOptions_Silent, 
                            "", 
                            0, 
                            0);
                    }
                    else if (file.EndsWith(".SLDASM", StringComparison.OrdinalIgnoreCase))
                    {
                        swModel = (ModelDoc2)swApp.OpenDoc6(
                            file, 
                            (int)swDocumentTypes_e.swDocASSEMBLY, 
                            (int)swOpenDocOptions_e.swOpenDocOptions_Silent, 
                            "", 
                            0, 
                            0);
                    }
                    else
                    {continue;
                    }

                    try
                    {
                       

                        if (swModel != null)
                        {
                            swModel.Visible = true;
                            // 转换为 DWG
                            asm2step.run(swModel);
                            
                            // 关闭已处理的文档
                            swApp.CloseDoc(swModel.GetTitle());
                                
            
            
                            Console.WriteLine($"已转换：{swModel.GetTitle()}");
                        }
                        else
                        {
                            Console.WriteLine($"无法打开文件：{file}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"转换失败 {file}: {ex.Message}");
                    }
                }
            }

        }
        
        
    }
  
}
    


