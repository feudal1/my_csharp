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
        public static void 标折弯尺寸(ISldWorks swApp)
        {
            try
            {
                var swModel = (ModelDoc2)swApp.ActiveDoc;
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
                double offset = 0.05;

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
                                if (ProcessBend(swModel, mathUtils, xform, xformData, bounds, subFeat, ref offset))
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

        static bool ProcessBend(ModelDoc2 swModel, MathUtility mathUtils, MathTransform xform,
            double[] xformData, double[] bounds, Feature bendFeat, ref double offset)
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
            var secondaryFaces = new List<Face>(); // 二级面或三级面
            foreach (var face in secondLevelFaces)
            {
                var s = (Surface)face.GetSurface();
                if (s.IsCylinder())
                {
                    Console.WriteLine("二级面为圆柱面，继续获取第三级面");
                    // 二级面是圆柱面，获取与轴线平行的边的相邻面
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
                                    if (adjFace != face && !secondaryFaces.Contains(adjFace))
                                    {
                                        secondaryFaces.Add(adjFace);
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    // 不是圆柱面，直接使用
                    secondaryFaces.Add(face);
                }
            }

            // 6. 找一级面与二级/三级面中互相平行的面作为标注面
            Face f1 = null, f2 = null;
            for (int i = 0; i < firstLevelFaces.Count; i++)
            {
                var s1 = (Surface)firstLevelFaces[i].GetSurface();
                if (!s1.IsPlane()) continue;
                var p1 = (double[])s1.PlaneParams;
                var n1 = new[] { p1[0], p1[1], p1[2] };

                for (int j = 0; j < secondaryFaces.Count; j++)
                {
                    var s2 = (Surface)secondaryFaces[j].GetSurface();
                    if (!s2.IsPlane()) continue;
                    var p2 = (double[])s2.PlaneParams;
                    var n2 = new[] { p2[0], p2[1], p2[2] };

                    // 检查两面是否平行（法向量平行）
                    if (IsParallel(n1, n2))
                    {
                        f1 = firstLevelFaces[i];
                        f2 = secondaryFaces[j];
                        Console.WriteLine("找到互相平行的标注面（一级面与二级/三级面）");
                        break;
                    }
                }
                if (f1 != null) break;
            }
            if (f1 == null || f2 == null)
            {
                Console.WriteLine("未找到互相平行的一级面与二级/三级面");
                return false;
            }

            // 打印选中面的信息
            var ts1 = (Surface)f1.GetSurface();
            var ts2 = (Surface)f2.GetSurface();
            Console.WriteLine($"标注面1（一级面）: 面积 = {f1.GetArea() * 1000000:F2} mm^2");
            Console.WriteLine($"标注面2（二级/三级面）: 面积 = {f2.GetArea() * 1000000:F2} mm^2");

            // 确定放置方式
            string placement = GetPlacement(f1, xformData);
            if (placement == "none") return false;

            // 标注
            SelectFaceEdge(swModel, f1);
            SelectFaceEdge(swModel, f2);

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

            return true;
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

        static void SelectFaceEdge(ModelDoc2 model, Face face)
        {
            var ent = (Entity)face;
            ent.Select4(true, null);
        }

        static double[] TransformPoint(MathUtility math, MathTransform xf, double[] pt)
        {
            var mp = (MathPoint)math.CreatePoint(pt);
            mp = (MathPoint)mp.MultiplyTransform(xf);
            return (double[])mp.ArrayData;
        }
    }
}
