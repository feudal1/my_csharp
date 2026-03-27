using System;
using tools;
using cad_tools;

namespace tools
{
    partial class Program
    {
        [Command("open_cad_doc_by_name", Description = "按名称打开 CAD 文档", Parameters = "<文档名称>", Group = "cad")]
        static void OpenCadDocByNameCommand(string[] args)
        {
            if (args.Length > 1)
            {
                open_cad_doc_by_name.run(args[1]);
            }
            else
            {
                Console.WriteLine("请提供文档名称参数");
            }
        }

        [Command("dxf2dwg", Description = "DXF 转换为 DWG 格式", Parameters = "<DXF 文件路径>", Group = "cad")]
        static void Dxf2DwgCommand(string[] args)
        {
            if (args.Length > 1)
            {
                dxf2dwg.run(args[1]);
            }
            else
            {
                Console.WriteLine("请提供 DXF 文件路径参数");
            }
        }

        [Command("folder2dwg", Description = "批量转换文件夹内 DXF 为 DWG", Parameters ="无", Group = "cad")]
        static void Folder2DwgCommand(string[] args)
        {
            var files = FolderPicker.GetFileNamesFromSelectedFolder();
            if (files != null)
            {
                foreach (var file in files)
                {
                    dxf2dwg.run(file);
                }
            }
        }
        [Command("folder2dxf", Description = "批量转换文件夹内 DXF 为 DWG", Parameters ="无", Group = "cad")]
        static void Folder2DXF	(string[] args)
        {
            var files = FolderPicker.GetFileNamesFromSelectedFolder();
            if (files != null)
            {
                foreach (var file in files)
                {
                    dwg2dxf.run(file);
                }
            }
        }



        [Command("mergedwg", Description = "遍历（子）文件夹：合并 DWG 并绘制边界框", Parameters = "无", Group = "cad")]
        static void FolderWithSubfoldersDrawDividerCommand(string[] args)
        {
            draw_divider.process_subfolders_with_divider();

        }
         [Command("get_all_dim_style", Description = "获取所有标注样式并添加 UUID4 后缀", Parameters = "无", Group = "cad")]
        static void GetAllDimStyleCommand(string[] args)
        {
            get_all_dim_style.run();
        }

    }
}

