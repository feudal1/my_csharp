namespace tools;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
public class one2step
{
    static public int run( ModelDoc2 swModel)
    {
        string fullPath = swModel.GetPathName();
        string partname = Path.GetFileNameWithoutExtension(fullPath);
        
        if (string.IsNullOrEmpty(fullPath))
        {
            Console.WriteLine("错误：文档尚未保存，请先保存文件。");
            return 0;
        }
                
        // 获取当前文件所在目录
        string? currentDirectory = Path.GetDirectoryName(fullPath);
        // 获取父文件夹
        string? parentDirectory = Directory.GetParent(currentDirectory)?.FullName;
        // 获取当前文件夹名称作为文件名
        string folderName = new DirectoryInfo(currentDirectory).Name;
                
        // 构建完整的输出路径
        string outputPath = Path.Combine(currentDirectory, $"{partname}.STEP");
                
        var result=swModel.SaveAs3(outputPath, 0, 2);
                
        Console.WriteLine($"{result}，已导出 STEP 文件到：{outputPath}");
        return 1 ;
    }

  
}