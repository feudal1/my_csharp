using System;
using SolidWorks.Interop.sldworks;
using tools;

namespace recognize.train_use
{
    /// <summary>
    /// WL 图核算法测试示例
    /// </summary>
    public class test_wl_graph_kernel
    {
        /// <summary>
        /// 测试当前打开的零件
        /// </summary>
        public static void RunTest(ISldWorks swApp, ModelDoc2 swModel)
        {
            if (swModel == null)
            {
                Console.WriteLine("错误：请先打开一个零件文档");
                return;
            }

            Console.WriteLine("\n========== WL 图核测试 ==========\n");

            // 构建单个零件的图
            PartGraph graph = FaceGraphBuilder.BuildGraph(swModel);
            
            if (graph == null || graph.Nodes.Count == 0)
            {
                Console.WriteLine("无法构建图结构");
                return;
            }

            // 打印节点信息
            Console.WriteLine($"\n图结构信息:");
            Console.WriteLine($"  零件名称：{graph.PartName}");
            Console.WriteLine($"  节点数量：{graph.Nodes.Count}");
            
            // 统计面类型分布
            var faceTypeStats = new System.Collections.Generic.Dictionary<string, int>();
            foreach (var node in graph.Nodes)
            {
                if (!faceTypeStats.ContainsKey(node.FaceType))
                {
                    faceTypeStats[node.FaceType] = 0;
                }
                faceTypeStats[node.FaceType]++;
            }

            Console.WriteLine($"\n  面类型分布:");
            foreach (var kvp in faceTypeStats)
            {
                Console.WriteLine($"    {kvp.Key,-10}: {kvp.Value} 个");
            }

            // 执行 WL 迭代
            Console.WriteLine($"\n执行 WL 迭代...");
            var freqList = WLGraphKernel.PerformWLIterations(graph, iterations: 1);

            // 显示每次迭代的标签频率
            for (int i = 0; i < freqList.Count; i++)
            {
                Console.WriteLine($"\n  迭代 {i}:");
                foreach (var kvp in freqList[i])
                {
                    Console.WriteLine($"    {kvp.Key}: {kvp.Value} 次");
                }
            }

            Console.WriteLine("\n====================================\n");
        }

        /// <summary>
        /// 比较两个零件的相似度
        /// </summary>
        public static void CompareTwoParts(ISldWorks swApp, ModelDoc2 model1, ModelDoc2 model2)
        {
            if (model1 == null || model2 == null)
            {
                Console.WriteLine("错误：需要两个有效的零件文档");
                return;
            }

            Console.WriteLine("\n========== 零件相似度对比 ==========\n");

            // 构建两个零件的图
            PartGraph graph1 = FaceGraphBuilder.BuildGraph(model1);
            PartGraph graph2 = FaceGraphBuilder.BuildGraph(model2);

            if (graph1 == null || graph2 == null)
            {
                Console.WriteLine("无法构建图结构");
                return;
            }

            // 计算相似度
            double similarity = SimilarityCalculator.CalculatePairSimilarity(graph1, graph2, iterations: 3);

            // 判断结果
            string conclusion;
            if (similarity > 0.8)
                conclusion = "高度相似 - 可能是同一系列零件";
            else if (similarity > 0.5)
                conclusion = "中等相似 - 具有类似的拓扑结构";
            else
                conclusion = "差异较大 - 拓扑结构明显不同";

            Console.WriteLine($"\n结论：{conclusion}");
            Console.WriteLine("\n======================================\n");
        }

        /// <summary>
        /// 批量分析文件夹中的零件
        /// </summary>
        public static void BatchAnalysis(ISldWorks swApp, string folderPath)
        {
            Console.WriteLine("\n========== 批量零件相似度分析 ==========\n");

            // 从文件夹加载零件
            var graphs = SimilarityCalculator.LoadPartsFromFolder(folderPath, swApp);

            if (graphs.Count < 2)
            {
                Console.WriteLine("零件数量不足，至少需要 2 个零件");
                return;
            }

            // 运行完整分析
            SimilarityCalculator.RunAnalysis(graphs, wlIterations: 1, decayFactor: 0.5);
        }
    }
}
