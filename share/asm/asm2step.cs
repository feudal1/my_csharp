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
        
        // 构建完整的输出路径
        string outputPath;
        
        // 如果零件名称包含"方管"，导出到特殊文件夹
        if (partname.Contains("方管"))
        {
            outputPath = Path.Combine(currentDirectory, "方管_step", $"{partname}.STEP");
        }
        else
        {
            outputPath = Path.Combine(currentDirectory, "step", $"{partname}.STEP");
        }
        
        string outputDirectory = Path.GetDirectoryName(outputPath);
        if (!Directory.Exists(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }       
        var result=swModel.SaveAs3(outputPath, 0, 2);
                
        Console.WriteLine($"{result}，已导出 STEP 文件到：{outputPath}");
        return 1 ;
    }

  
}