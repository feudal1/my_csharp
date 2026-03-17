//https://help.solidworks.com/2023/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IFlatPatternFeatureData_members.html
//https://help.solidworks.com/2023/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IFeature~GetTypeName.html
//https://help.solidworks.com/2023/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISheetMetalFeatureData_members.html
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
namespace tools
{
    public class exportdwg2_body
    {
        private static string? directory = "";
        private static PartDoc swPart = null;
        private  static string fullPath = "";
        private static string partname = "";
        private  static double[] dataAlignment=new double[12];
        private  static int options;
        private static int successcount = 0;
        
                static public void exportfeature(Body2 body)
                {
                    string outputfile = "";
                    string thickness = "无";
                    object[] features = (object[])body.GetFeatures();
                    foreach ( object objFeature in features)
                    {
                        
                        Feature swFeature = (Feature)objFeature;
                   
                        if (swFeature.GetTypeName2() == "SheetMetal")
                    {
                        SheetMetalFeatureData swSheetMetalData = (SheetMetalFeatureData)swFeature.GetDefinition();
                        thickness =  Math.Round(swSheetMetalData.Thickness*1000,2).ToString();
                         outputfile = directory + "\\"+"出图"+"\\" + "下料" + "\\" + thickness;
                        if (!Directory.Exists(outputfile))
                        {
                            Directory.CreateDirectory(outputfile);
                        }
                    }

                    if (swFeature.GetTypeName2() == "FlatPattern")
                    {
                         string dwgFileName = outputfile + "\\" +  partname+"_"+swFeature.Name + ".dwg";
                        var swFlatPatt = (FlatPatternFeatureData)swFeature.GetDefinition();
                         swFeature .SetSuppression((int)swFeatureSuppressionAction_e.swUnSuppressFeature);
                         var fixface = (    Face2)swFlatPatt.FixedFace2;
                        var fixface_area=Math.Round( fixface.GetArea()*1000000,2);
                        
                        var subfeat = (Feature)swFeature.GetFirstSubFeature();
                        bool hasbend=false;
                        while (subfeat != null)
                        {

                            Console.WriteLine(subfeat.GetTypeName());
                            if (subfeat.GetTypeName() == "UiBend")
                            {
                                hasbend=true;

                            }
                            subfeat = (Feature)subfeat.GetNextSubFeature();
                        }
                        Face2 bigger_face = fixface;
                        if (!hasbend) {
                        
                            var edgeobjs = (object[])fixface.GetEdges();
                            foreach (var edgeobj in edgeobjs)
                            {
                                var edge = (Edge)edgeobj;
                                var twoface=(object[])edge.GetTwoAdjacentFaces();
                                var other_face=((Face2)twoface[0]).IsSame(fixface) ? (Face2)twoface[1] : (Face2)twoface[0];
                                var other_face_area=Math.Round( other_face.GetArea()*1000000,2);
                                bigger_face=other_face_area > fixface_area ? other_face : fixface;
                                



                            }

                        }
                   
                          
                        var bigger_face_ent = (Entity)bigger_face;
                      
                        var select_result=bigger_face_ent.Select4(true, null);

                        if (select_result)
                        {Console.WriteLine($"选择面成功");
                        }

                        var success=swPart.ExportToDWG(dwgFileName, fullPath, (int)swExportToDWG_e.swExportToDWG_ExportSelectedFacesOrLoops, true, dataAlignment, false, false, options, null);
swFeature .SetSuppression((int)swFeatureSuppressionAction_e.swSuppressFeature);
                        Console.WriteLine($"{success},导出{dwgFileName}");
                        if (success) { successcount++; }

                        
                  
               
                    }
                      
                    }
                  
                 
                    
                       
                
              
        }
        static public int run(ModelDoc2 swModel)
        {
           
            try
            {
                successcount = 0;
                              fullPath = swModel.GetPathName();
                              partname = Path.GetFileNameWithoutExtension(fullPath);

                if (string.IsNullOrEmpty(fullPath))
                {
                    Console.WriteLine("错误：文档尚未保存，请先保存文件。");
                    
                }
                 directory = Path.GetDirectoryName(fullPath);
                if (string.IsNullOrEmpty(directory))
                {
                    Console.WriteLine("错误：无法获取文件所在目录。");
                  
                }
                swPart = (PartDoc)swModel;
    
               
         
                dataAlignment[0] = 0.0;
                dataAlignment[1] = 0.0;
                dataAlignment[2] = 0.0;
                dataAlignment[3] = 1.0;
                dataAlignment[4] = 0.0;
                dataAlignment[5] = 0.0;
                dataAlignment[6] = 0.0;
                dataAlignment[7] = 1.0;
                dataAlignment[8] = 0.0;
                dataAlignment[9] = 1.0;
                dataAlignment[10] = 0.0;
                dataAlignment[11] = 0.0;
                 options = 97;
Console.OutputEncoding = Encoding.UTF8;
               

                var partdoc = (PartDoc)swModel;
                var bodys = (object[])partdoc.GetBodies2((int)swBodyType_e.swSolidBody, false);

                foreach (var objbody in bodys)
                {
                    var body = (Body2)objbody;
                   
                   
                        exportfeature(body);
                    
                  
                    
                }
                
                
                

            }
            catch (Exception ex)
            {
                Console.WriteLine($"发生错误: {ex.Message}");
                Console.WriteLine("提示：请确保 SolidWorks 正在运行。");
            }
        
            return successcount;
        }
    }
}