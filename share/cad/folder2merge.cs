namespace cad_tools;

public class folder2merge
{
    public static void run()
    { 
        var files = FolderPicker.GetFileNamesFromSelectedFolder();
        if (files != null)
        {
            double currentX = 0.0;
            double partSpacing = 10.0; // 零件之间的间距
                
            foreach (var file in files)
            {
                Console.WriteLine($"准备插入文件：{System.IO.Path.GetFileName(file)}, 位置=({currentX:F2}, 0.00)");
                    
                var maxPoint = merge_dwg.run(file, currentX, 0.0);
                    
                if (maxPoint != null)
                {
                    double partMaxX = maxPoint[0];
                    currentX = partMaxX + partSpacing;
                    Console.WriteLine($"已插入，下一个位置 startX={currentX:F2}");
                }
            }
        }
    }
}