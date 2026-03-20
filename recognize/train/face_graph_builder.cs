using System;
using System.Collections.Generic;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace tools
{
    /// <summary>
    /// WL 图核算法中的面图节点
    /// </summary>
    public class FaceNode
    {
        public int Id { get; set; }                      // 节点索引
        public Face2 FaceObject { get; set; }            // 面对象引用
        public string FaceType { get; set; }             // 面类型标签
        public double Area { get; set; }                 // 面积
        public List<int> NeighborIds { get; set; }       // 邻居节点索引列表
        public string CurrentLabel { get; set; }         // 当前迭代标签
        
        public FaceNode()
        {
            NeighborIds = new List<int>();
        }
    }

    /// <summary>
    /// 零件的面邻接图表征
    /// </summary>
    public class PartGraph
    {
        public string PartName { get; set; }             // 零件名称
        public List<FaceNode> Nodes { get; set; }        // 节点列表
        public Dictionary<string, int> LabelFrequency { get; set; }  // 标签频率统计
        
        public PartGraph()
        {
            Nodes = new List<FaceNode>();
            LabelFrequency = new Dictionary<string, int>();
        }
    }

    /// <summary>
    /// 从 SolidWorks 零件构建面邻接图
    /// </summary>
    public static class FaceGraphBuilder
    {
        /// <summary>
        /// 从零件文档构建面邻接图
        /// </summary>
        public static PartGraph BuildGraph(ModelDoc2 swModel, string partName = "")
        {
            if (swModel == null)
            {
                Console.WriteLine("错误：没有打开的活动文档。");
                return null;
            }

            PartDoc partDoc = (PartDoc)swModel;
            PartGraph graph = new PartGraph();
            graph.PartName = string.IsNullOrEmpty(partName) ? swModel.GetTitle() : partName;

            // 存储邻接关系的字典：面索引 -> 邻居面索引集合
            var adjacencyDict = new Dictionary<int, HashSet<int>>();
            var faceIndexMap = new Dictionary<Face2, int>();
            var facesList = new List<Face2>();
            int faceCount = 0;

            // 获取所有实体
            object[] vBodies = (object[])partDoc.GetBodies2((int)swBodyType_e.swSolidBody, false);
            
            if (vBodies == null || vBodies.Length == 0)
            {
                Console.WriteLine("警告：未找到实体 body");
                return graph;
            }

            // 遍历所有 body，提取面和邻接关系
            foreach (Body2 body in vBodies)
            {
                object[] vEdges = (object[])body.GetEdges();
                if (vEdges == null) continue;

                foreach (Edge edge in vEdges)
                {
                    var twoAdjacentFaces = (object[])edge.GetTwoAdjacentFaces();
                    if (twoAdjacentFaces == null || twoAdjacentFaces.Length < 2) continue;

                    Face2 face1 = (Face2)twoAdjacentFaces[0];
                    Face2 face2 = (Face2)twoAdjacentFaces[1];

                    // 为每个面分配唯一索引
                    if (!faceIndexMap.ContainsKey(face1))
                    {
                        faceIndexMap[face1] = faceCount++;
                        facesList.Add(face1);
                    }

                    if (!faceIndexMap.ContainsKey(face2))
                    {
                        faceIndexMap[face2] = faceCount++;
                        facesList.Add(face2);
                    }

                    int index1 = faceIndexMap[face1];
                    int index2 = faceIndexMap[face2];

                    // 初始化邻接矩阵条目
                    if (!adjacencyDict.ContainsKey(index1))
                        adjacencyDict[index1] = new HashSet<int>();

                    if (!adjacencyDict.ContainsKey(index2))
                        adjacencyDict[index2] = new HashSet<int>();

                    // 添加邻接关系
                    adjacencyDict[index1].Add(index2);
                    adjacencyDict[index2].Add(index1);
                }
            }

            // 创建节点并分配初始标签
            for (int i = 0; i < faceCount; i++)
            {
                Face2 face = facesList[i];
                Surface surface = face.IGetSurface();
                
                // 确定面类型
                string faceType = GetFaceType(surface);
                
                // 计算面积 (mm²)
                double area = Math.Round(face.GetArea() * 1000000, 2);

                FaceNode node = new FaceNode
                {
                    Id = i,
                    FaceObject = face,
                    FaceType = faceType,
                    Area = area,
                    CurrentLabel = faceType,  // 初始标签为面类型
                    NeighborIds = adjacencyDict.ContainsKey(i) 
                        ? adjacencyDict[i].ToList() 
                        : new List<int>()
                };

                graph.Nodes.Add(node);
            }

            Console.WriteLine($"零件 [{graph.PartName}] 构建完成：{faceCount} 个面");
            return graph;
        }

        /// <summary>
        /// 根据曲面几何特性确定面类型
        /// </summary>
        private static string GetFaceType(Surface surface)
        {
            if (surface.IsPlane()) return "平面";
            if (surface.IsCylinder()) return "圆柱面";
            if (surface.IsCone()) return "圆锥面";
            if (surface.IsSphere()) return "球面";
            if (surface.IsTorus()) return "圆环面";
     
            return "其他曲面";
        }
    }
}
