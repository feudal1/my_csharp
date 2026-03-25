using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using tools;
using cad_tools;

namespace tools
{
    partial class Program
    {
        // ========== 拓扑标注相关命令 ==========
        
        [Command("label", Description = "标注当前零件的拓扑图并输入标签", Parameters = "无", Group = "train")]
        static void LabelPart(string[] args)
        {
            if (Program.SwModel == null)
            {
                Console.WriteLine("错误：请先打开一个零件文档");
                return;
            }

            Console.WriteLine("=== 零件拓扑标注 ===\n");
            TopologyLabeler.LabelCurrentPart(Program.SwModel, wlIterations: 1);
        }

        [Command("label_quick", Description = "快速标注当前零件（仅计算存储，不立即标注）", Parameters = "无", Group = "train")]
        static void QuickLabelPart(string[] args)
        {
            if (Program.SwModel == null)
            {
                Console.WriteLine("错误：请先打开一个零件文档");
                return;
            }

            TopologyLabelingExample.QuickLabel(Program.SwApp!, Program.SwModel);
        }

        [Command("view_parts", Description = "查看所有已标注的零件", Parameters = "无", Group = "train")]
        static void ViewAllParts(string[] args)
        {
            TopologyLabeler.ViewAllParts();
        }

        [Command("label_search", Description = "按标注类别搜索零件", Parameters = "[类别名称] [可选的值]", Group = "train")]
        static void SearchLabels(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("用法：label_search [类别名称] [可选的值]");
                Console.WriteLine("示例：label_search 结构类型 框架");
                return;
            }

            string category = args[0];
            string? value = args.Length > 1 ? args[1] : null;

            TopologyLabeler.SearchByCategory(category, value);
        }

        [Command("label_stats", Description = "显示数据库统计信息，打印一下数据库所有数据", Parameters = "无", Group = "train")]
        static void ShowLabelStats(string[] args)
        {
            TopologyLabeler.ShowStatistics();
        }

        [Command("label_batch", Description = "批量标注文件夹中的所有零件", Parameters = "[文件夹路径]", Group = "train")]
        static void BatchLabel(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("用法：label_batch [文件夹路径]");
                Console.WriteLine("示例：label_batch E:\\parts");
                return;
            }

            string folderPath = args[0];

            if (!System.IO.Directory.Exists(folderPath))
            {
                Console.WriteLine($"错误：文件夹不存在：{folderPath}");
                return;
            }

            if (Program.SwApp == null)
            {
                Console.WriteLine("错误：无法连接到 SolidWorks");
                return;
            }

            TopologyLabelingExample.BatchProcessExample(Program.SwApp, folderPath);
        }

        [Command("label_export", Description = "导出零件的 WL 结果为 JSON", Parameters = "[零件 ID]", Group = "train")]
        static void ExportWL(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("用法：label_export [零件 ID]");
                Console.WriteLine("提示：使用 view_parts 查看零件 ID");
                return;
            }

            if (int.TryParse(args[0], out int partId))
            {
                TopologyLabeler.ExportPartWL(partId);
            }
            else
            {
                Console.WriteLine("错误：零件 ID 必须是数字");
            }
        }

        [Command("label_delete", Description = "删除指定的标注 (删除零件 id 为 x 的数据)", Parameters = "[标注 ID]", Group = "train")]
        static void DeleteLabel(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("用法：label_delete [标注 ID]");
                Console.WriteLine("提示：使用 label_search 或 view_parts 查看标注 ID");
                return;
            }
        
            if (int.TryParse(args[0], out int labelId))
            {
                TopologyLabeler.DeleteLabel(labelId);
            }
            else
            {
                Console.WriteLine("错误：标注 ID 必须是数字");
            }
        }
        
        [Command("delete_part", Description = "删除零件 ID 为 x 的所有数据（包括 WL 结果和标注）", Parameters = "[零件 ID]", Group = "train")]
        static void DeletePart(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("用法：delete_part [零件 ID]");
                Console.WriteLine("提示：使用 view_parts 查看零件 ID");
                return;
            }
        
            if (int.TryParse(args[0], out int partId))
            {
                TopologyLabeler.DeletePartData(partId);
            }
            else
            {
                Console.WriteLine("错误：零件 ID 必须是数字");
            }
        }
    }
}