using tools;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace recognize
{
    class MAIN
    {
    [STAThread]
        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            
            // 无参数 - 显示帮助
            if (args.Length == 0)
            {
                ShowHelp();
                return;
            }

            // 连接到 SolidWorks
            var swApp = Connect.run();
            if (swApp == null)
            {
                Console.WriteLine("错误：无法连接到 SolidWorks 应用程序。");
                return;
            }

            var swModel = (ModelDoc2)swApp.ActiveDoc;

            // 解析命令
            string command = args[0].ToLower();

            try
            {
                switch (command)
                {
                    case "edges":
                    case "edge":
                        // 获取所有边信息
                        get_all_edges.run(swModel);
                        break;

                    case "faces":
                    case "face":
                        // 获取所有面信息
                        get_all_face.run(swModel);
                        break;

                    case "dims":
                    case "dim":
                        // 获取尺寸信息
                        get_dim_info.run(swModel);
                        break;

                    case "views":
                    case "view":
                        // 获取视图图结构
                        get_views_graph.run(swModel);
                        break;

                    case "tap":
                    case "taphole":
                        // 获取 tap hole 信息
                        get_tap_hole.run( swModel);
                        break;

                    // Train 相关命令
                    case "train-test":
                    case "test":
                        // 测试当前打开的零件
                        if (swModel == null)
                        {
                            Console.WriteLine("错误：请先打开一个零件文档");
                            return;
                        }
                        recognize.train_use.test_wl_graph_kernel.RunTest(swApp, swModel);
                        break;

                    case "train-compare":
                    case "compare":
                        // 比较两个零件
                        if (args.Length < 3)
                        {
                            Console.WriteLine("错误：compare 命令需要两个零件文件路径");
                            Console.WriteLine("用法：recognize compare <文件 1> <文件 2>");
                            return;
                        }
                        
                        string file1 = args[1];
                        string file2 = args[2];
                        
                        if (!File.Exists(file1) || !File.Exists(file2))
                        {
                            Console.WriteLine("错误：文件不存在");
                            return;
                        }
                        
                        // 打开两个零件
                        int errors = 0, warnings = 0;
                        ModelDoc2 model1 = swApp.OpenDoc6(
                            file1, 
                            (int)swDocumentTypes_e.swDocPART, 
                            (int)swOpenDocOptions_e.swOpenDocOptions_Silent, 
                            "", 
                            ref errors, 
                            ref warnings);
                        
                        ModelDoc2 model2 = swApp.OpenDoc6(
                            file2, 
                            (int)swDocumentTypes_e.swDocPART, 
                            (int)swOpenDocOptions_e.swOpenDocOptions_Silent, 
                            "", 
                            ref errors, 
                            ref warnings);
                        
                        if (model1 != null && model2 != null)
                        {
                            recognize.train_use.test_wl_graph_kernel.CompareTwoParts(swApp, model1, model2);
                            
                            // 关闭文档
                            swApp.CloseDoc(model1.GetTitle());
                            swApp.CloseDoc(model2.GetTitle());
                        }
                        else
                        {
                            Console.WriteLine("错误：无法打开零件文件");
                        }
                        break;

                    case "train-batch":
                    case "batch":
                        // 批量分析文件夹中的零件 - 使用文件夹选择对话框
                        using (var folderDialog = new System.Windows.Forms.FolderBrowserDialog())
                        {
                            folderDialog.Description = "请选择包含 SolidWorks 零件文件的文件夹";
                            folderDialog.ShowNewFolderButton = false;
                            
                            if (folderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                            {
                                string folderPath = folderDialog.SelectedPath;
                                recognize.train_use.test_wl_graph_kernel.BatchAnalysis(swApp, folderPath);
                            }
                            else
                            {
                                Console.WriteLine("已取消文件夹选择");
                            }
                        }
                        break;

            


                    case "help":
                    case "-h":
                    case "--help":
                        ShowHelp();
                        break;

                    default:
                        Console.WriteLine($"未知命令：{command}");
                        ShowHelp();
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n执行命令时出错：{ex.Message}");
                Console.WriteLine($"详细信息：{ex.StackTrace}");
            }
        }

        /// <summary>
        /// 显示帮助信息
        /// </summary>
        static void ShowHelp()
        {
            Console.WriteLine("\n===== Recognize - SolidWorks 特征识别工具 =====\n");
            Console.WriteLine("用法：recognize [命令] [参数]\n");
            Console.WriteLine("可用命令:");
            Console.WriteLine("  edges              获取所有边信息");
            Console.WriteLine("  faces              获取所有面信息");
            Console.WriteLine("  dims               获取尺寸信息");
            Console.WriteLine("  views              获取视图图结构");
            Console.WriteLine("  tap                获取 tap hole 信息");
            Console.WriteLine("\nTrain 相关命令:");
            Console.WriteLine("  test               测试当前打开的零件 (WL 图核)");
            Console.WriteLine("  compare <文件 1> <文件 2>  比较两个零件的相似度");
            Console.WriteLine("  batch              批量分析文件夹中的所有零件 (弹窗选择文件夹)");
            Console.WriteLine("\n装配体 WL 批量处理命令:");
            Console.WriteLine("  asm-wl             基于装配体的 WL 批量处理 (构建所有零件的面邻接图)");
            Console.WriteLine("  asm-wl-similar     计算装配体中所有零件的相似度矩阵");
            Console.WriteLine("  asm-wl-example     运行基础使用示例");
            Console.WriteLine("  asm-wl-advanced    运行高级示例 (多次迭代对比)");
            Console.WriteLine("\n示例:");
            Console.WriteLine("  recognize edges");
            Console.WriteLine("  recognize test");
            Console.WriteLine("  recognize compare C:\\parts\\part1.sldprt C:\\parts\\part2.sldprt");
            Console.WriteLine("  recognize batch      (会弹出文件夹选择对话框)");
            Console.WriteLine();
        }
    }
}