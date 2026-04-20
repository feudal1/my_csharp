using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using View = SolidWorks.Interop.sldworks.View;

namespace tools
{
    public class benddim
    {
        public static void AddBendDimensions(ISldWorks swApp)
        {
            try
            {
                var swModel = (ModelDoc2)swApp.ActiveDoc;
                swApp.CommandInProgress = true;
                if (swModel == null) { Console.WriteLine("没有活动文档"); return; }

                var swSelMgr = (SelectionMgr)swModel.SelectionManager;
                if (swSelMgr.GetSelectedObjectType3(1, -1) != (int)swSelectType_e.swSelDRAWINGVIEWS)
                { Console.WriteLine("请先选择一个视图"); return; }

                var view = (View)swSelMgr.GetSelectedObject(1);
                var partDoc = (PartDoc)view.ReferencedDocument;
                if (partDoc == null) { Console.WriteLine("无法获取零件文档"); return; }

                var xform = (MathTransform)view.ModelToViewTransform;
                var xformData = (double[])xform.ArrayData;
                var bounds = (double[])view.GetOutline();
                var mathUtils = swApp.IGetMathUtility();

                int count = 0;
                double offset = 0.005;
                
                // 全局已标注面组合集合，用于跨折弯去重
                var dimensionedPairs = new HashSet<string>();

                // 遍历 Body 的特征找折弯
                foreach (Body2 body in (object[])partDoc.GetBodies2((int)swBodyType_e.swSolidBody, false))
                {
                    foreach (Feature feat in (object[])body.GetFeatures())
                    {
                        var subFeat = (Feature)feat.GetFirstSubFeature();
                        while (subFeat != null)
                        {
                            if (subFeat.GetTypeName() == "OneBend" )
                            {
                                Console.WriteLine("找到折弯特征" + subFeat.Name);
                                if (ProcessBend(swModel, view, mathUtils, xform, xformData, bounds, subFeat, ref offset, dimensionedPairs))
                                    count++;
                            }

                            subFeat = (Feature)subFeat.GetNextSubFeature();
                        }
                    }
                }

                Console.WriteLine($"标注完成，共 {count} 个折弯");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"错误: {ex.Message}");
            }
        }

        static bool ProcessBend(ModelDoc2 swModel, View view, MathUtility mathUtils, MathTransform xform,
            double[] xformData, double[] bounds, Feature bendFeat, ref double offset, HashSet<string> dimensionedPairs)
        {
            // 1. 找最大圆柱面
            Face cylFace = null;
            double[] axis = null;
            double[] center = null;
            double maxArea = 0;

            foreach (Face f in (object[])bendFeat.GetFaces())
            {
                var s = (Surface)f.GetSurface();
                if (s.IsCylinder())
                {
                    double area = f.GetArea();
                    if (area > maxArea)
                    {
                        maxArea = area;
                        var p = (double[])s.CylinderParams;
                        axis = new[] { p[3], p[4], p[5] };
                        center = new[] { p[0], p[1], p[2] };
                        cylFace = f;
                    }
                }
            }
            if (cylFace == null) return false;
            Console.WriteLine("最大圆柱面面积: " + maxArea * 1000000 + " mm^2");

            // 2. 找圆柱面的edge里和圆柱轴线平行的edge（直边）
            var parallelEdges = new List<Edge>();
            foreach (Edge e in (object[])cylFace.GetEdges())
            {
                var c = (Curve)e.GetCurve();
                if (c.IsLine())
                {
                    var lineParams = (double[])c.LineParams;
                    var edgeDir = new[] { lineParams[3], lineParams[4], lineParams[5] };
                    if (IsParallel(edgeDir, axis))
                    {
                        parallelEdges.Add(e);
                    }
                }
            }
            if (parallelEdges.Count == 0) return false;
            Console.WriteLine("找到 " + parallelEdges.Count + " 条与轴线平行的边");

            // 3. 用edge的相交面找下一级面
            var firstLevelFaces = new List<Face>();
            foreach (var edge in parallelEdges)
            {
                foreach (Face f in (object[])edge.GetTwoAdjacentFaces())
                {
                    if (f != cylFace && !firstLevelFaces.Contains(f))
                    {
                        firstLevelFaces.Add(f);
                    }
                }
            }
            if (firstLevelFaces.Count == 0) return false;
            Console.WriteLine("找到 " + firstLevelFaces.Count + " 个下一级面");

            // 4. 在下一级面里找与圆柱轴平行的最远的线，然后取相交面
            var secondLevelFaces = new List<Face>();
            foreach (var face in firstLevelFaces)
            {
                Edge farthestEdge = null;
                double maxDist = -1;

                foreach (Edge e in (object[])face.GetEdges())
                {
                    var c = (Curve)e.GetCurve();
                    if (c.IsLine())
                    {
                        var lineParams = (double[])c.LineParams;
                        var edgeDir = new[] { lineParams[3], lineParams[4], lineParams[5] };
                        if (IsParallel(edgeDir, axis))
                        {
                            // 计算边到圆柱中心的距离
                            var pt = new[] { lineParams[0], lineParams[1], lineParams[2] };
                            double dist = PointToAxisDistance(pt, center, axis);
                            if (dist > maxDist)
                            {
                                maxDist = dist;
                                farthestEdge = e;
                            }
                        }
                    }
                }

                if (farthestEdge != null)
                {
                    foreach (Face f in (object[])farthestEdge.GetTwoAdjacentFaces())
                    {
                        if (f != face && !secondLevelFaces.Contains(f))
                        {
                            secondLevelFaces.Add(f);
                        }
                    }
                }
            }
            if (secondLevelFaces.Count < 2) return false;
            Console.WriteLine("找到 " + secondLevelFaces.Count + " 个二级面");

            // 5. 检查二级面是否为圆柱面，如果是则获取第三级面
            var secondaryFaces = new List<(Face face, string level)>(); // 面及其级别
            foreach (var face in secondLevelFaces)
            {
                var s = (Surface)face.GetSurface();
                if (s.IsCylinder())
                {
                    Console.WriteLine("二级面为圆柱面，继续获取第三级面");
                    // 二级面是圆柱面，获取与轴线平行的边的相邻面作为三级面
                    int thirdLevelCount = 0;
                    foreach (Edge e in (object[])face.GetEdges())
                    {
                        var c = (Curve)e.GetCurve();
                        if (c.IsLine())
                        {
                            var lineParams = (double[])c.LineParams;
                            var edgeDir = new[] { lineParams[3], lineParams[4], lineParams[5] };
                            if (IsParallel(edgeDir, axis))
                            {
                                foreach (Face adjFace in (object[])e.GetTwoAdjacentFaces())
                                {
                                    // 排除：自身、一级面、折弯圆柱面
                                    if (adjFace == face || firstLevelFaces.Contains(adjFace) || adjFace == cylFace) continue;
                                    if (!secondaryFaces.Any(x => x.face == adjFace))
                                    {
                                        secondaryFaces.Add((adjFace, "三级面"));
                                        thirdLevelCount++;
                                    }
                                }
                            }
                        }
                    }
                    Console.WriteLine($"  找到 {thirdLevelCount} 个三级面");
                }
                else
                {
                    // 不是圆柱面，直接使用
                    secondaryFaces.Add((face, "二级面"));
                }
            }

            // 6. 为每个一级面找到对应的平行二级/三级面并标注
            // 使用全局已标注面集合防止重复标注（跨折弯共享）
            int dimensionCount = 0;
            
            foreach (var firstFace in firstLevelFaces)
            {
                var s1 = (Surface)firstFace.GetSurface();
                if (!s1.IsPlane()) continue;
                var p1 = (double[])s1.PlaneParams;
                var n1 = new[] { p1[0], p1[1], p1[2] };

                // 为当前一级面找平行的二级/三级面
                Face matchedSecondFace = null;
                string matchedLevel = "";
                foreach (var item in secondaryFaces)
                {
                    var secFace = item.face;
                    var s2 = (Surface)secFace.GetSurface();
                    if (!s2.IsPlane()) continue;
                    var p2 = (double[])s2.PlaneParams;
                    var n2 = new[] { p2[0], p2[1], p2[2] };

                    // 检查两面是否平行（法向量平行）且不是同一个面
                    if (IsParallel(n1, n2) && secFace != firstFace)
                    {
                        // 生成唯一键：两面法向量、面积和配对级别的组合
                        var pairKey = GetFacePairKey(firstFace, secFace, item.level);
                        if (dimensionedPairs.Contains(pairKey))
                        {
                            Console.WriteLine($"跳过已标注的平行面组合：一级面面积 = {firstFace.GetArea() * 1000000:F2} mm^2, {item.level}面积 = {secFace.GetArea() * 1000000:F2} mm^2");
                            continue;
                        }
                        
                        matchedSecondFace = secFace;
                        matchedLevel = item.level;
                        dimensionedPairs.Add(pairKey);
                        Console.WriteLine($"找到一对平行面：一级面面积 = {firstFace.GetArea() * 1000000:F2} mm^2, {matchedLevel}面积 = {secFace.GetArea() * 1000000:F2} mm^2");
                        break;
                    }
                }

                if (matchedSecondFace == null) continue;

                // 确定放置方式
                string placement = GetPlacement(firstFace, xformData);
                if (placement == "none") continue;

                // 标注 - 获取3D边并在视图中查找对应可见边
                // 获取所有候选边（按长度降序），逐个尝试找到有可见边的
                var edgeCandidates1 = GetEdgesForDimension(firstFace, axis, "一级面");
                var edgeCandidates2 = GetEdgesForDimension(matchedSecondFace, axis, matchedLevel);
                
                Edge visEdge1 = null, visEdge2 = null;
                
                foreach (var e1 in edgeCandidates1)
                {
                    visEdge1 = FindVisibleEdge(view, e1);
                    if (visEdge1 != null) break;
                }
                
                if (visEdge1 == null)
                {
                    Console.WriteLine("一级面的标注边在视图中均不可见，跳过此对");
                    continue;
                }
                
                foreach (var e2 in edgeCandidates2)
                {
                    visEdge2 = FindVisibleEdge(view, e2);
                    if (visEdge2 != null) break;
                }
                
                if (visEdge2 == null)
                {
                    Console.WriteLine("二级面的标注边在视图中均不可见，跳过此对");
                    continue;
                }

                // 创建 SelectData 并设置视图上下文
                var selMgr = (SelectionMgr)swModel.SelectionManager;
                var selData = selMgr.CreateSelectData();
                selData.View = view;

                // 将可见边转换为 Entity 并选择
                ((Entity)visEdge1).Select4(true, selData);
                ((Entity)visEdge2).Select4(true, selData);

                double x, y;
                if (placement == "h")
                {
                    x = (bounds[0] + bounds[2]) / 2;
                    y = bounds[3] - offset;
                }
                else
                {
                    y = (bounds[1] + bounds[3]) / 2;
                    x = bounds[0] + offset;
                }
                swModel.AddDimension2(x, y, 0);
                offset += 0.005;
                swModel.ClearSelection2(true);
                dimensionCount++;
            }

            return dimensionCount > 0;
        }

        static bool IsParallel(double[] a, double[] b)
        {
            double dot = Math.Abs(a[0] * b[0] + a[1] * b[1] + a[2] * b[2]);
            double ma = Math.Sqrt(a[0] * a[0] + a[1] * a[1] + a[2] * a[2]);
            double mb = Math.Sqrt(b[0] * b[0] + b[1] * b[1] + b[2] * b[2]);
            return Math.Abs(dot - ma * mb) < 0.001;
        }

        static double PointToAxisDistance(double[] point, double[] axisCenter, double[] axisDir)
        {
            // 计算点到轴线的距离
            // 向量从轴线中心指向点
            double[] v = { point[0] - axisCenter[0], point[1] - axisCenter[1], point[2] - axisCenter[2] };
            
            // 投影到轴线方向
            double dot = v[0] * axisDir[0] + v[1] * axisDir[1] + v[2] * axisDir[2];
            double axisLen = Math.Sqrt(axisDir[0] * axisDir[0] + axisDir[1] * axisDir[1] + axisDir[2] * axisDir[2]);
            double proj = dot / axisLen;
            
            // 投影点
            double[] projPoint = {
                axisCenter[0] + proj * axisDir[0] / axisLen,
                axisCenter[1] + proj * axisDir[1] / axisLen,
                axisCenter[2] + proj * axisDir[2] / axisLen
            };
            
            // 距离
            double dx = point[0] - projPoint[0];
            double dy = point[1] - projPoint[1];
            double dz = point[2] - projPoint[2];
            return Math.Sqrt(dx * dx + dy * dy + dz * dz);
        }

        static string GetPlacement(Face face, double[] xf)
        {
            var s = (Surface)face.GetSurface();
            if (!s.IsPlane()) return "none";
            var p = (double[])s.PlaneParams;
            double nx = Math.Abs(p[0]), ny = Math.Abs(p[1]), nz = Math.Abs(p[2]);

            if (nx > ny && nx > nz)
                return Math.Round(xf[0]) != 0 ? "h" : (Math.Round(xf[1]) != 0 ? "v" : "none");
            if (ny > nx && ny > nz)
                return Math.Round(xf[3]) != 0 ? "h" : (Math.Round(xf[4]) != 0 ? "v" : "none");
            if (nz > nx && nz > ny)
                return Math.Round(xf[6]) != 0 ? "h" : (Math.Round(xf[7]) != 0 ? "v" : "none");
            return "none";
        }

        /// <summary>
        /// 获取面上所有不与轴线平行的直边，按长度降序排列（用于尺寸标注）
        /// 折弯标注需要选择不平行于折弯轴线的边，返回多条候选边供视图中逐个匹配
        /// </summary>
        static List<Edge> GetEdgesForDimension(Face face, double[] axis, string faceLevel = "")
        {
            var candidates = new List<(Edge edge, double len)>();

            foreach (Edge e in (object[])face.GetEdges())
            {
                var c = (Curve)e.GetCurve();
                if (c.IsLine())
                {
                    var lineParams = (double[])c.LineParams;
                    var edgeDir = new[] { lineParams[3], lineParams[4], lineParams[5] };
                    
                    // 跳过与轴线平行的边，折弯标注需要不平行的边
                    if (IsParallel(edgeDir, axis))
                        continue;

                    double startParam = 0, endParam = 0;
                    bool isClosed = false, isPeriodic = false;
                    c.GetEndParams(out startParam, out endParam, out isClosed, out isPeriodic);
                    double len = Math.Abs(endParam - startParam);
                    
                    candidates.Add((e, len));
                }
            }

            // 按长度降序排列，优先使用长边
            candidates.Sort((a, b) => b.len.CompareTo(a.len));
            
            var label = string.IsNullOrEmpty(faceLevel) ? "" : $"[{faceLevel}] ";
            foreach (var item in candidates)
                Console.WriteLine($"  {label}候选标注边，长度: {item.len * 1000:F2} mm");
            
            if (candidates.Count == 0)
                Console.WriteLine($"  {label}未找到不平行于轴线的直边");

            return candidates.Select(x => x.edge).ToList();
        }

        /// <summary>
        /// 在视图中查找3D边对应的可见边
        /// 按组件获取可见边（返回可转换为Edge的对象），通过3D端点坐标近似匹配解决COM封送问题
        /// 注意：IDrawingEdge接口在SolidWorks C#互操作中不存在，不能类型转换；
        ///       GetVisibleEntities(null,...) 返回的对象也是__ComObject无法转换；
        ///       正确做法是用 GetVisibleEntities(comp,...) 按组件获取，返回的对象可转换为Edge
        /// </summary>
        static Edge FindVisibleEdge(View view, Edge modelEdge)
        {
            // 获取目标3D边的端点坐标
            var startVert = (Vertex)modelEdge.GetStartVertex();
            var endVert = (Vertex)modelEdge.GetEndVertex();
            if (startVert == null || endVert == null) return null;

            var edgeStart = (double[])startVert.GetPoint();
            var edgeEnd = (double[])endVert.GetPoint();
            if (edgeStart == null || edgeEnd == null) return null;

            // 获取视图中的可见组件
            var visibleComps = (object[])view.GetVisibleComponents();
            if (visibleComps == null) return null;

            foreach (Component2 comp in visibleComps)

            {
                if (comp == null) continue;

                // 按组件获取可见边（与 GetVisibleEntities(null,...) 不同，按组件获取返回可转换为Edge的对象）
                var visibleEdges = (object[])view.GetVisibleEntities(
                    comp, (int)swViewEntityType_e.swViewEntityType_Edge);
                if (visibleEdges == null) continue;

                foreach (object obj in visibleEdges)
                {
                    if (obj == null) continue;

                    // GetVisibleEntities(comp,...) 返回的对象可转换为 Edge
                    if (!(obj is Edge visEdge)) continue;

                    try
                    {
                        // 先用引用判断（速度快）
                        if (Object.ReferenceEquals(visEdge, modelEdge)) return visEdge;

                        // 再用3D端点坐标近似判断（解决COM封送导致的引用不一致问题）
                        var meStartVert = (Vertex)visEdge.GetStartVertex();
                        var meEndVert = (Vertex)visEdge.GetEndVertex();
                        if (meStartVert == null || meEndVert == null) continue;

                        var meStart = (double[])meStartVert.GetPoint();
                        var meEnd = (double[])meEndVert.GetPoint();
                        if (meStart == null || meEnd == null) continue;

                        bool coordsMatch =
                            (IsPointClose(edgeStart, meStart) && IsPointClose(edgeEnd, meEnd))
                            ||
                            (IsPointClose(edgeStart, meEnd) && IsPointClose(edgeEnd, meStart));

                        if (coordsMatch) return visEdge;
                    }
                    catch
                    {
                        continue;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// 判断两个3D点是否在容差范围内近似相等
        /// </summary>
        static bool IsPointClose(double[] p1, double[] p2, double tol = 0.001)
        {
            if (p1 == null || p2 == null || p1.Length < 3 || p2.Length < 3) return false;
            var dx = Math.Abs(p1[0] - p2[0]);
            var dy = Math.Abs(p1[1] - p2[1]);
            var dz = Math.Abs(p1[2] - p2[2]);
            var dist = Math.Sqrt(dx * dx + dy * dy + dz * dz);
            var close = dx < tol && dy < tol && dz < tol;
           // Console.WriteLine($"[IsPointClose] p1=({p1[0]:F6},{p1[1]:F6},{p1[2]:F6}) p2=({p2[0]:F6},{p2[1]:F6},{p2[2]:F6}) dx={dx:F6} dy={dy:F6} dz={dz:F6} dist={dist:F6} tol={tol} => {close}");
            return close;
        }

        static double[] TransformPoint(MathUtility math, MathTransform xf, double[] pt)
        {
            var mp = (MathPoint)math.CreatePoint(pt);
            mp = (MathPoint)mp.MultiplyTransform(xf);
            return (double[])mp.ArrayData;
        }

        /// <summary>
        /// 生成一对面的唯一标识键，用于防止重复标注
        /// 基于Face对象的引用生成唯一键，确保只有完全相同的物理面才会被去重
        /// </summary>
        static string GetFacePairKey(Face face1, Face face2, string level)
        {
            // 使用Face对象的HashCode生成唯一标识
            // 按HashCode排序生成键，确保 (A,B) 和 (B,A) 生成相同的键
            int hash1 = face1.GetHashCode();
            int hash2 = face2.GetHashCode();
            
            var hashes = new[] { hash1, hash2 };
            Array.Sort(hashes);
            
            // 加入配对级别，确保一级对二级 与 一级对三级 不会互相去重
            return $"{hashes[0]}|{hashes[1]}|{level}";
        }
    }
}
