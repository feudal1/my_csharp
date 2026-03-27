using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using tools;
using cad_tools;

namespace tools
{
    partial class Program
    {
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


        [Command("select_part_by_name", Description = "按名称选择零件", Parameters = "零件名称", Group = "solidworks")]
        static void SelectPartByNameCommand(string[] args)
        {
            if (swApp == null || swModel == null) return;

           
           select_part_byname.run(swModel, args[0]);
            
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

        [Command("asm2bom", Description = "装配体导出 bom", Parameters = "无", Group = "solidworks")]
        static void Asm2bom(string[] args)
        {
            if (swApp == null || swModel == null) return;

           
                asm2bom.run(swApp,swModel);
       
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
                    {
                        continue;
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
