using System;
using SolidWorks.Interop.sldworks;

namespace tools
{
    /// <summary>
    /// 拓扑标注系统快速入门示例
    /// </summary>
    public static class 拓扑标注快速入门
    {
        /// <summary>
        /// 运行快速入门示例
        /// </summary>
        public static void Run(SldWorks swApp, ModelDoc2 swModel)
        {
            Console.WriteLine("========================================");
            Console.WriteLine("   零件拓扑标注系统 - 快速入门");
            Console.WriteLine("========================================\n");

            // 步骤 1: 初始化数据库
            Console.WriteLine("【步骤 1】初始化数据库...");
            TopologyLabeler.Initialize();
            Console.WriteLine();

            // 步骤 2: 标注当前零件
            Console.WriteLine("【步骤 2】标注当前零件...");
            if (swModel != null)
            {
                Console.WriteLine($"当前零件：{swModel.GetTitle()}\n");
                TopologyLabeler.LabelCurrentPart(swModel, wlIterations: 1);
            }
            else
            {
                Console.WriteLine("× 请先打开一个零件文档");
            }

            Console.WriteLine();

            // 步骤 3: 查看统计信息
            Console.WriteLine("【步骤 3】查看数据库统计...");
            TopologyLabeler.ShowStatistics();

            Console.WriteLine();
            Console.WriteLine("========================================");
            Console.WriteLine("   入门完成！");
            Console.WriteLine("========================================");
            Console.WriteLine("\n提示命令:");
            Console.WriteLine("  label          - 标注新零件");
            Console.WriteLine("  view_parts     - 查看所有零件");
            Console.WriteLine("  label_search   - 搜索零件");
            Console.WriteLine("  label_stats    - 查看统计");
            Console.WriteLine("  label_batch    - 批量处理文件夹");
            Console.WriteLine("  label_export   - 导出 WL 结果");
            Console.WriteLine("\n详细说明请查看：拓扑标注系统使用说明.md");
        }

        /// <summary>
        /// 演示完整的标注工作流
        /// </summary>
        public static void DemoWorkflow(SldWorks swApp)
        {
            Console.WriteLine("\n=== 标注工作流演示 ===\n");

            // 1. 准备阶段
            Console.WriteLine("1. 准备阶段:");
            Console.WriteLine("   - 在 SolidWorks 中打开要标注的零件");
            Console.WriteLine("   - 确保零件已保存（有完整路径）");
            Console.WriteLine();

            // 2. 自动计算
            Console.WriteLine("2. 自动计算拓扑特征:");
            Console.WriteLine("   使用 label_quick 命令快速存储特征");
            Console.WriteLine("   系统会自动:");
            Console.WriteLine("     • 提取零件的所有面");
            Console.WriteLine("     • 构建面邻接图");
            Console.WriteLine("     • 执行 WL 迭代计算");
            Console.WriteLine("     • 存储到 SQLite 数据库");
            Console.WriteLine();

            // 3. 手动标注
            Console.WriteLine("3. 添加语义标注:");
            Console.WriteLine("   使用 label 命令进入交互式标注:");
            Console.WriteLine("     • 输入标注类别（如'结构类型'）");
            Console.WriteLine("     • 输入标注值（如'框架'）");
            Console.WriteLine("     • 设置置信度（可选）");
            Console.WriteLine("     • 添加备注说明（可选）");
            Console.WriteLine();

            // 4. 查询使用
            Console.WriteLine("4. 查询与检索:");
            Console.WriteLine("   使用 label_search 按类别查找:");
            Console.WriteLine("   例：label_search 结构类型 框架");
            Console.WriteLine();

            // 5. 批量处理
            Console.WriteLine("5. 批量处理:");
            Console.WriteLine("   使用 label_batch 处理整个文件夹:");
            Console.WriteLine("   例：label_batch E:\\projects\\parts");
            Console.WriteLine();

            Console.WriteLine("=== 演示结束 ===\n");
        }
    }
}
