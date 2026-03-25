using System;
using SolidWorks.Interop.sldworks;
using tools;

namespace recognize.train_use
{
    /// <summary>
    /// WL 图核使用示例
    /// </summary>
    public class wl_usage_example
    {
        /// <summary>
        /// 示例 1: 分析当前打开的零件
        /// </summary>
        public static void Example1_SinglePart(ISldWorks swApp, ModelDoc2 swModel)
        {
            // 构建图结构
            PartGraph graph = FaceGraphBuilder.BuildGraph(swModel);
            
            // 执行 WL 迭代
            var freqList = WLGraphKernel.PerformWLIterations(graph, iterations: 3);
            
            // 查看每次迭代的标签频率
            for (int i = 0; i < freqList.Count; i++)
            {
                Console.WriteLine($"迭代 {i} 的标签分布:");
                foreach (var kvp in freqList[i])
                {
                    Console.WriteLine($"  标签 {kvp.Key}: {kvp.Value} 次");
                }
            }
        }

        /// <summary>
        /// 示例 2: 比较两个零件
        /// </summary>
        public static void Example2_CompareTwoParts(ISldWorks swApp, ModelDoc2 part1, ModelDoc2 part2)
        {
            // 构建两个零件的图
            PartGraph graph1 = FaceGraphBuilder.BuildGraph(part1);
            PartGraph graph2 = FaceGraphBuilder.BuildGraph(part2);
            
            // 计算相似度
            double similarity = SimilarityCalculator.CalculatePairSimilarity(graph1, graph2, iterations: 3);
            
            Console.WriteLine($"相似度：{similarity * 100:F2}%");
            
            if (similarity > 0.8)
                Console.WriteLine("高度相似");
            else if (similarity > 0.5)
                Console.WriteLine("中等相似");
            else
                Console.WriteLine("差异较大");
        }

        /// <summary>
        /// 示例 3: 批量分析文件夹中的所有零件
        /// </summary>
        public static void Example3_BatchAnalysis(ISldWorks swApp, string folderPath)
        {
            // 从文件夹加载所有零件
            var graphs = SimilarityCalculator.LoadPartsFromFolder(folderPath, swApp);
            
            // 运行完整分析，生成相似度矩阵
            SimilarityCalculator.RunAnalysis(graphs, wlIterations: 3, decayFactor: 0.5);
        }

        /// <summary>
        /// 示例 4: 自定义 WL 参数
        /// </summary>
        public static void Example4_CustomParameters(ModelDoc2 part1, ModelDoc2 part2)
        {
            PartGraph graph1 = FaceGraphBuilder.BuildGraph(part1);
            PartGraph graph2 = FaceGraphBuilder.BuildGraph(part2);
            
            // 执行更多次 WL 迭代 (捕捉更复杂的结构)
            var freqList1 = WLGraphKernel.PerformWLIterations(graph1, iterations: 5);
            var freqList2 = WLGraphKernel.PerformWLIterations(graph2, iterations: 5);
            
            // 使用不同的权重衰减因子
            // decayFactor=0.3 表示更重视早期迭代 (局部结构)
            // decayFactor=0.8 表示更重视后期迭代 (全局结构)
            double sim_fastDecay = WLGraphKernel.CalculateComprehensiveSimilarity(
                freqList1, freqList2, decayFactor: 0.3);
            
            double sim_slowDecay = WLGraphKernel.CalculateComprehensiveSimilarity(
                freqList1, freqList2, decayFactor: 0.8);
            
            Console.WriteLine($"快速衰减相似度 (侧重局部): {sim_fastDecay * 100:F2}%");
            Console.WriteLine($"慢速衰减相似度 (侧重全局): {sim_slowDecay * 100:F2}%");
        }

        /// <summary>
        /// 示例 5: 手动计算特定迭代的相似度
        /// </summary>
        public static void Example5_ManualCalculation(ModelDoc2 part1, ModelDoc2 part2)
        {
            PartGraph graph1 = FaceGraphBuilder.BuildGraph(part1);
            PartGraph graph2 = FaceGraphBuilder.BuildGraph(part2);
            
            var freqList1 = WLGraphKernel.PerformWLIterations(graph1, iterations: 3);
            var freqList2 = WLGraphKernel.PerformWLIterations(graph2, iterations: 3);
            
            // 单独查看某次迭代的相似度
            // 迭代 0: 仅基于面类型
            double sim_iter0 = WLGraphKernel.CalculateSimilarity(freqList1[0], freqList2[0]);
            
            // 迭代 1: 考虑一阶邻域
            double sim_iter1 = WLGraphKernel.CalculateSimilarity(freqList1[1], freqList2[1]);
            
            // 迭代 2: 考虑二阶邻域
            double sim_iter2 = WLGraphKernel.CalculateSimilarity(freqList1[2], freqList2[2]);
            
            Console.WriteLine($"迭代 0 相似度 (仅面类型): {sim_iter0 * 100:F2}%");
            Console.WriteLine($"迭代 1 相似度 (一阶邻域): {sim_iter1 * 100:F2}%");
            Console.WriteLine($"迭代 2 相似度 (二阶邻域): {sim_iter2 * 100:F2}%");
        }
    }
}
