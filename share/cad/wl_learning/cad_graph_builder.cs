using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;

namespace cad_tools
{
    /// <summary>
    /// CAD图形构建器 - 从CAD文档提取2D图形特征
    /// </summary>
    public static class CADGraphEdgeBuilder
    {
        /// <summary>
        /// 从当前CAD文档构建图形
        /// </summary>
        /// <param name="acadDoc">CAD文档</param>
        /// <param name="graphName">图形名称（可选，默认为文件名）</param>
        /// <returns>CAD 2D图</returns>
        public static CADGraphEdgeGraph BuildGraphFromDocument(AcadDocument acadDoc, string graphName = "")
        {
            var graph = new CADGraphEdgeGraph
            {
                SourceFile = acadDoc.FullName,
                GraphName = string.IsNullOrEmpty(graphName) 
                    ? System.IO.Path.GetFileNameWithoutExtension(acadDoc.FullName) 
                    : graphName
            };

            try
            {
                // 获取模型空间中的所有实体
                var modelSpace = acadDoc.ModelSpace;
                int nodeId = 0;
                
                // 第一遍：创建所有节点
                var nodeMap = new Dictionary<object, CADGraphEdgeNode>();
                
                foreach (AcadEntity entity in modelSpace)
                {
                    var node = CreateNodeFromEntity(entity, nodeId++);
                    if (node != null)
                    {
                        graph.Nodes.Add(node);
                        nodeMap[entity] = node;
                    }
                }

                // 第二遍：建立连接关系（基于端点重合）
                BuildConnections(graph, nodeMap);

                Console.WriteLine($"✓ 成功构建CAD图形：{graph.GraphName}，包含 {graph.Nodes.Count} 个节点");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"× 构建CAD图形失败：{ex.Message}");
            }

            return graph;
        }

        /// <summary>
        /// 从CAD实体创建节点
        /// </summary>
        private static CADGraphEdgeNode CreateNodeFromEntity(AcadEntity entity, int nodeId)
        {
            var node = new CADGraphEdgeNode
            {
                Id = nodeId
            };

            try
            {
                switch (entity)
                {
                    case AcadLine line:
                        node.EdgeType = "Line";
                        node.GeometryValue = CalculateLineLength(line);
                        var startPt = (double[])line.StartPoint;
                        var endPt = (double[])line.EndPoint;
                        node.IsHorizontal = Math.Abs(endPt[1] - startPt[1]) < 0.001;
                        node.IsVertical = Math.Abs(endPt[0] - startPt[0]) < 0.001;
                        node.Angle = CalculateLineAngle(line);
                        break;

                    case AcadArc arc:
                        node.EdgeType = "Arc";
                        node.GeometryValue = arc.Radius;
                        node.Angle = (arc.EndAngle - arc.StartAngle) * 180.0 / Math.PI;
                        break;

                    case AcadCircle circle:
                        node.EdgeType = "Circle";
                        node.GeometryValue = circle.Radius;
                        node.IsHorizontal = true;
                        node.IsVertical = true;
                        break;

                    case AcadLWPolyline polyline:
                        node.EdgeType = "Polyline";
                        node.GeometryValue = CalculatePolylineLength(polyline);
                        break;

                    case AcadEllipse ellipse:
                        node.EdgeType = "Ellipse";
                        node.GeometryValue = ellipse.MajorRadius;
                        break;

                    default:
                        return null; // 跳过不支持的实体类型
                }

                node.CurrentLabel = node.EdgeType;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"警告：无法解析实体：{ex.Message}");
                return null;
            }

            return node;
        }

        /// <summary>
        /// 建立连接关系（基于端点重合）
        /// </summary>
        private static void BuildConnections(CADGraphEdgeGraph graph, Dictionary<object, CADGraphEdgeNode> nodeMap)
        {
            double tolerance = 0.001; // 容差

            foreach (var nodeA in graph.Nodes)
            {
                foreach (var nodeB in graph.Nodes)
                {
                    if (nodeA.Id == nodeB.Id) continue;

                    if (AreConnected(nodeA, nodeB, tolerance))
                    {
                        if (!nodeA.ConnectedNodes.Contains(nodeB.Id))
                        {
                            nodeA.ConnectedNodes.Add(nodeB.Id);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 判断两个节点是否连接（需要具体实现端点检测）
        /// </summary>
        private static bool AreConnected(CADGraphEdgeNode nodeA, CADGraphEdgeNode nodeB, double tolerance)
        {
            // TODO: 实现端点重合检测逻辑
            // 这里需要根据实体类型获取端点坐标
            return false;
        }

        /// <summary>
        /// 计算直线长度
        /// </summary>
        private static double CalculateLineLength(AcadLine line)
        {
            var startPt = (double[])line.StartPoint;
            var endPt = (double[])line.EndPoint;
            double dx = endPt[0] - startPt[0];
            double dy = endPt[1] - startPt[1];
            double dz = endPt[2] - startPt[2];
            return Math.Sqrt(dx * dx + dy * dy + dz * dz);
        }

        /// <summary>
        /// 计算直线角度（相对于X轴）
        /// </summary>
        private static double CalculateLineAngle(AcadLine line)
        {
            var startPt = (double[])line.StartPoint;
            var endPt = (double[])line.EndPoint;
            double dx = endPt[0] - startPt[0];
            double dy = endPt[1] - startPt[1];
            double angle = Math.Atan2(dy, dx) * 180.0 / Math.PI;
            return angle < 0 ? angle + 360 : angle;
        }

        /// <summary>
        /// 计算多段线长度
        /// </summary>
        private static double CalculatePolylineLength(AcadLWPolyline polyline)
        {
            double length = 0;
            var coords = (double[])polyline.Coordinates;
            
            for (int i = 0; i < coords.Length - 2; i += 2)
            {
                double dx = coords[i + 2] - coords[i];
                double dy = coords[i + 3] - coords[i + 1];
                length += Math.Sqrt(dx * dx + dy * dy);
            }
            
            return length;
        }

        /// <summary>
        /// 批量构建文件夹中所有CAD文档的图形
        /// </summary>
        public static List<CADGraphEdgeGraph> BuildGraphsFromFolder(string folderPath)
        {
            var graphs = new List<CADGraphEdgeGraph>();
            
            if (!System.IO.Directory.Exists(folderPath))
            {
                Console.WriteLine($"错误：文件夹不存在：{folderPath}");
                return graphs;
            }

            var dwgFiles = System.IO.Directory.GetFiles(folderPath, "*.dwg");
            Console.WriteLine($"\n找到 {dwgFiles.Length} 个DWG文件");

            var acadApp = CadConnect.GetOrCreateInstance();
            if (acadApp == null)
            {
                Console.WriteLine("错误：无法连接到AutoCAD");
                return graphs;
            }

            foreach (var dwgFile in dwgFiles)
            {
                try
                {
                    Console.WriteLine($"\n处理：{System.IO.Path.GetFileName(dwgFile)}");
                    var doc = acadApp.Documents.Open(dwgFile, false, true);
                    var graph = BuildGraphFromDocument(doc, System.IO.Path.GetFileNameWithoutExtension(dwgFile));
                    graphs.Add(graph);
                    doc.Close(false);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"× 处理失败：{ex.Message}");
                }
            }

            return graphs;
        }
    }
}
