using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using tools;
using cad_tools;

namespace tools
{
    partial class Program
    {
        [Command("new_drw", Description = "新建工程图", Parameters = "无", Group = "solidworks")]
        static void NewDrwCommand(string[] args)
        {
            if (swApp == null || swModel == null) return;
            add_name2info.run(swModel);
            New_drw.run(swApp, swModel);
        }

        [Command("drw2dwg", Description = "将工程图转换为 DWG 格式", Parameters = "无", Group = "solidworks")]
        static void Drw2DwgCommand(string[] args)
        {
            if (swApp == null || swModel == null) return;

            var dwgFileName = drw2dwg.run(swModel, swApp);
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
 [Command("analyze_face", Description = "分析当前选中的面（面积、类型、法向等）", Parameters = "需先在 SolidWorks 中选择面", Group = "solidworks")]
        static void AnalyzeFaceCommand(string[] args)
        {
            if (swModel == null) return;
                
            select_face_recognize.run(swModel);
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
            Console.WriteLine($"成功导出：{pngPath}");

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
                await vlmService.ChatAsync( prompt, pngPath);
                
                Console.WriteLine("\n----------------------------");
                Console.WriteLine("✅ 分析完成。");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n❌ VLM 调用失败：{ex.Message}");
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



        [Command("dimension_bends", Description = "自动标注选中视图的折弯尺寸", Parameters = "需先选择工程图视图", Group = "solidworks")]
        static void DimensionBendsCommand(string[] args)
        {
            if (swApp == null) return;

            benddim.标折弯尺寸(swApp);
        }
    }
}
