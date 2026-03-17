using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using System;

namespace cad_tools
{
    public class draw_divider
    {

        /// <summary>
        /// 处理子文件夹遍历并为每个子文件夹的 DWG 文件绘制边界框（支持有子文件夹和没有子文件夹两种情况）
        /// </summary>
        static public void process_subfolders_with_divider()
        {
            string? selectedFolder = FolderPicker.SelectFolder();
            if (string.IsNullOrEmpty(selectedFolder))
            {
                Console.WriteLine("未选择文件夹，操作取消。");
                return;
            }
           
            double partSpacing = 15.0; // 零件之间的间距
            double folderSpacing = 100.0; // 文件夹之间的间距
            
            // 获取所有子文件夹
            var subDirectories = Directory.GetDirectories(selectedFolder);
            
            // 情况 1：如果有子文件夹，遍历子文件夹
            if (subDirectories.Length > 0)
            {
                double currentFolderMaxY = 0.0; // 当前文件夹的最大 Y 值
                
                foreach (var subDir in subDirectories)
                {
                    Console.WriteLine($"\n处理子文件夹：{Path.GetFileName(subDir)}");
                    
                    // 获取子文件夹中的所有 DWG 文件
                    var dwgFiles = Directory.GetFiles(subDir, "*.dwg");
                    
                    if (dwgFiles.Length > 0)
                    {
                        AcadApplication? acadApp = CadConnect.GetOrCreateInstance();
                        
                        if (acadApp != null && acadApp.ActiveDocument != null)
                        {
                            double folderStartX = 0.0; // 每个文件夹从 X=0 开始
                            double folderMinY = currentFolderMaxY + folderSpacing; // 上一个文件夹的 maxY + 间距
                            double folderCurrentMaxY = folderMinY; // 当前文件夹的 maxY 初始值
                            
                            Console.WriteLine($"开始合并子文件夹 {Path.GetFileName(subDir)} 的 {dwgFiles.Length} 个文件，起始 Y={folderMinY:F2}");
                            
                            // 合并该子文件夹的所有 DWG，每个文件按序号递增 X 偏移
                            for (int i = 0; i < dwgFiles.Length; i++)
                            {
                                Console.WriteLine($"  准备插入第 {i+1} 个文件：{Path.GetFileName(dwgFiles[i])}, 位置=({folderStartX:F2}, {folderMinY:F2})");
                                var maxPoint = merge_dwg.run(dwgFiles[i], folderStartX, folderMinY);
                                
                                if (maxPoint != null)
                                {
                                    // maxPoint: [maxX, maxY]
                                    double partMaxX = maxPoint[0];
                                    double partMaxY = maxPoint[1];
                                    
                                    Console.WriteLine($"  返回尺寸：宽={partMaxX:F2}, 高={partMaxY:F2}");
                                    
                                    // 更新当前文件夹的 maxY
                                    if (partMaxY > folderCurrentMaxY)
                                    {
                                        folderCurrentMaxY = partMaxY;
                                    }
                                    
                                    // 下一个零件的 startX = 当前零件的 maxX + 间距
                                    folderStartX = partMaxX + partSpacing;
                                    Console.WriteLine($"  下一个位置 startX={folderStartX:F2}");
                                }
                            }
                            
                            // 更新全局的 currentFolderMaxY
                            if (folderCurrentMaxY > currentFolderMaxY)
                            {
                                currentFolderMaxY = folderCurrentMaxY;
                            }
                            
                        
                            
                            Console.WriteLine($"已绘制子文件夹 {Path.GetFileName(subDir)} 的边界框：X(0 到 {folderStartX:F2}), Y({folderMinY:F2} 到 {folderCurrentMaxY:F2})");
                        }
                        else
                        {
                            Console.WriteLine($"警告：无法获取 AutoCAD 文档，跳过边界框绘制");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"子文件夹 {Path.GetFileName(subDir)} 中没有 DWG 文件");
                    }
                }
            }
            // 情况 2：如果没有子文件夹，直接处理当前文件夹中的 DWG 文件
            else
            {
                Console.WriteLine($"\n未找到子文件夹，直接处理当前文件夹：{Path.GetFileName(selectedFolder)}");
                
                // 获取当前文件夹中的所有 DWG 文件
                var dwgFiles = Directory.GetFiles(selectedFolder, "*.dwg");
                
                if (dwgFiles.Length > 0)
                {
                    AcadApplication? acadApp = CadConnect.GetOrCreateInstance();
                    
                    if (acadApp != null && acadApp.ActiveDocument != null)
                    {
                        double folderStartX = 0.0;
                        double folderMinY = 0.0;
                        double folderCurrentMaxY = folderMinY;
                        
                        Console.WriteLine($"开始合并当前文件夹的 {dwgFiles.Length} 个文件");
                        
                        // 合并所有 DWG，每个文件按序号递增 X 偏移
                        for (int i = 0; i < dwgFiles.Length; i++)
                        {
                            var maxPoint = merge_dwg.run(dwgFiles[i], folderStartX, folderMinY);
                            
                            if (maxPoint != null)
                            {
                                double partMaxX = maxPoint[0];
                                double partMaxY = maxPoint[1];
                                
                                if (partMaxY > folderCurrentMaxY)
                                {
                                    folderCurrentMaxY = partMaxY;
                                }
                                
                                folderStartX = partMaxX + partSpacing;
                            }
                        }
                        
                  
                        
                        Console.WriteLine($"已绘制当前文件夹的边界框：X(0 到 {folderStartX:F2}), Y({folderMinY:F2} 到 {folderCurrentMaxY:F2})");
                    }
                    else
                    {
                        Console.WriteLine($"警告：无法获取 AutoCAD 文档，跳过边界框绘制");
                    }
                }
                else
                {
                    Console.WriteLine($"当前文件夹中没有 DWG 文件");
                }
            }
            
            Console.WriteLine("\n所有处理完成！");
        }

       
    }
}