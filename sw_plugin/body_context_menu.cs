using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorksTools;
using System;
using System.Diagnostics;
using SolidWorks.Interop.swconst;
using System.Collections.Generic;
using System.IO;
using tools;
namespace SolidWorksAddinStudy
{
    /// <summary>
    /// 右键菜单管理器 - 使用 AddMenuPopupItem4 为 FeatureManager 设计树中的实体添加右键菜单
    /// </summary>
  public partial class AddinStudy 
{

              public string open_dwg()
              {
                  ShowOutputWindow();
                  var acswModel = (ModelDoc2)swApp.ActiveDoc;
                  var (swModel, body) = get_select_body(acswModel);
                  open_select_dwg.run(swModel, (Body2)body);
                  
            return "a";

    }          
      
      public string new_drawing_from_part()
      {
          ShowOutputWindow();
          try
          {
              Debug.WriteLine("开始创建工程图");
              if (swApp == null)
              {
                  Debug.WriteLine("SolidWorks 未初始化");
                  return "SolidWorks 未初始化";
              }

              var acswModel = (ModelDoc2)swApp.ActiveDoc;
                  var (swModel, body) = get_select_body(acswModel);
              if (swModel == null)
              {
                  Debug.WriteLine("没有打开的文档");
                  swApp.SendMsgToUser("请先打开一个零件文档");
                  return "没有打开的文档";
              }

              // 添加名称到自定义信息
              add_name2info.run(swModel);
              
              // 创建新工程图
              New_drw.run(swApp, swModel);
              
              Debug.WriteLine("工程图已创建");
              return "工程图已创建";
          }
          catch (Exception ex)
          {
              Debug.WriteLine($"创建工程图失败：{ex.Message}");
              swApp?.SendMsgToUser($"创建工程图失败：{ex.Message}");
              return $"创建工程图失败：{ex.Message}";
          }
      }

       public (ModelDoc2,IBody2) get_select_body(ModelDoc2 swModel)
       {
           var swSelMgr = (SelectionMgr)swModel.SelectionManager;
            var selboj = swSelMgr.IGetSelectedObject(1);
            var face= (Face2)selboj;
            var body = (IBody2)face.GetBody();
            if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                var comp = (Component2)swSelMgr.GetSelectedObjectsComponent(1);
                swModel = (ModelDoc2)comp.GetModelDoc2();
            }

            Console.WriteLine(body.Name);
            return (swModel, body);
       }

       public string export_flat_pattern()
       {
           ShowOutputWindow();
           try
           {
               Debug.WriteLine("开始导出展开");
               if (swApp == null)
               {
                   Debug.WriteLine("SolidWorks 未初始化");
                   return "SolidWorks 未初始化";
               }

 var acswModel = (ModelDoc2)swApp.ActiveDoc;
                  var (swModel, body) = get_select_body(acswModel);
               if (swModel == null)
               {
                   Debug.WriteLine("没有打开的文档");
                   swApp.SendMsgToUser("请先打开一个零件文档");
                   return "没有打开的文档";
               }
               
               checkk_factor.run(swApp, swModel);
               exportdwg2_body.run(swModel);
               
               Debug.WriteLine("展开导出完成");
               return "展开导出完成";
           }
           catch (Exception ex)
           {
               Debug.WriteLine($"导出展开失败：{ex.Message}");
               swApp?.SendMsgToUser($"导出展开失败：{ex.Message}");
               return $"导出展开失败：{ex.Message}";
           }
       }

        public string copy_model_dwg_to_clipboard()
        {
            ShowOutputWindow();
            try
            {
                Debug.WriteLine("开始复制工程图DWG到剪贴板");
                if (swApp == null)
                {
                    Debug.WriteLine("SolidWorks 未初始化");
                    return "SolidWorks 未初始化";
                }

                var acswModel = (ModelDoc2)swApp.ActiveDoc;
                var (swModel, body) = get_select_body(acswModel);
                if (swModel == null)
                {
                    Debug.WriteLine("没有打开的文档");
                    swApp.SendMsgToUser("请先打开一个零件文档");
                    return "没有打开的文档";
                }

                // 获取模型路径
                string fullPath = swModel.GetPathName();
                if (string.IsNullOrEmpty(fullPath))
                {
                    Debug.WriteLine("错误：文档尚未保存，请先保存文件。");
                    swApp.SendMsgToUser("请先保存当前文档");
                    return "文档未保存";
                }

                string directory = Path.GetDirectoryName(fullPath);
                string partName = Path.GetFileNameWithoutExtension(fullPath);
                
                // 查找工程图文件
                string drwFileName = directory + "\\" + partName + ".SLDDRW";
                if (!File.Exists(drwFileName))
                {
                    Debug.WriteLine($"工程图文件不存在: {drwFileName}");
                    swApp.SendMsgToUser($"工程图文件不存在: {partName}.SLDDRW");
                    return $"工程图文件不存在";
                }

                // 打开工程图以获取DWG路径
                int errors = 0, warnings = 0;
                ModelDoc2 drwModel = (ModelDoc2)swApp.OpenDoc6(
                    drwFileName,
                    (int)swDocumentTypes_e.swDocDRAWING,
                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                    "",
                    ref errors,
                    ref warnings);

                if (drwModel == null)
                {
                    Debug.WriteLine($"无法打开工程图: {drwFileName}");
                    swApp.SendMsgToUser($"无法打开工程图");
                    return $"无法打开工程图";
                }

                // 使用 drw2dwg 的逻辑来获取DWG文件路径
                string dwgFileName = get_dwg_path_from_drawing(drwModel, swApp);
                
                // 关闭工程图
                swApp.CloseDoc(drwModel.GetTitle());

                // 检查DWG文件是否存在
                if (File.Exists(dwgFileName))
                {
                    // 使用Windows Forms的Clipboard类将文件复制到剪贴板
                    System.Collections.Specialized.StringCollection fileList = new System.Collections.Specialized.StringCollection();
                    fileList.Add(dwgFileName);
                    System.Windows.Forms.Clipboard.SetFileDropList(fileList);
                    
                    Debug.WriteLine($"已将工程图DWG文件复制到剪贴板: {dwgFileName}");
                    return $"已将DWG复制到剪贴板: {Path.GetFileName(dwgFileName)}";
                }
                else
                {
                    Debug.WriteLine($"DWG文件不存在: {dwgFileName}");
                    swApp.SendMsgToUser($"DWG文件不存在，请先转换工程图为DWG");
                    return $"DWG文件不存在";
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"复制DWG文件到剪贴板失败：{ex.Message}");
                swApp?.SendMsgToUser($"复制DWG文件到剪贴板失败：{ex.Message}");
                return $"复制DWG文件到剪贴板失败：{ex.Message}";
            }
        }

        private string get_dwg_path_from_drawing(ModelDoc2 drwModel, SldWorks swApp)
        {
            string fullpath = drwModel.GetPathName();
            string? directory = Path.GetDirectoryName(fullpath);
            
            var drawingDoc = (DrawingDoc)drwModel;
            var swSheet = (Sheet)drawingDoc.IGetCurrentSheet();
            var swViews = (object[])swSheet.GetViews();
            var partDoc = ((SolidWorks.Interop.sldworks.View)swViews[1]).ReferencedDocument;
            
            string outputfile;
            
            // 判断是否为装配体
            if (partDoc.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                outputfile = directory + "\\" + "出图" + "\\" + "焊接图";
            }
            else
            {
                var thickness = get_thickness.run(partDoc);
                
                if (thickness == 0)
                {
                    outputfile = directory + "\\" + "出图" + "\\" + "CNC";
                }
                else
                {
                    var meterial = ((PartDoc)partDoc).GetMaterialPropertyName(out _);
                    var meterialthick = thickness.ToString();
                    
                    if (meterial != null && meterial.ToLower().Contains("sus"))
                    {
                        meterialthick = "sus" + thickness.ToString();
                    }
                    
                    outputfile = directory + "\\" + "出图" + "\\" + "工程图" + "\\" + meterialthick;
                }
            }
            
            string dwgFileName = outputfile + "\\" + Path.GetFileNameWithoutExtension(fullpath) + ".dwg";
            return dwgFileName;
        }

        /// <summary>
        /// 初始化实体右键菜单
        /// </summary>
        public void PopupMenuInitialize()
        {
            try
            {
                if (swApp == null || iCmdMgr == null)
                {
                    Debug.WriteLine("SolidWorks 未初始化");
                    return;
                }

                
         
  
                 swApp.AddMenuPopupItem2((int)swDocumentTypes_e.swDocPART, addinCookieID, (int)swSelectType_e.swSelFACES, "export_dwg", "export_flat_pattern", "", "", "");
                 swApp.AddMenuPopupItem2((int)swDocumentTypes_e.swDocASSEMBLY, addinCookieID, (int)swSelectType_e.swSelFACES, "export_dwg", "export_flat_pattern", "", "", "");
                            swApp.AddMenuPopupItem2((int)swDocumentTypes_e.swDocPART, addinCookieID, (int)swSelectType_e.swSelFACES, "new_drw", "new_drawing_from_part", "", "", "");
                swApp.AddMenuPopupItem2((int)swDocumentTypes_e.swDocASSEMBLY, addinCookieID, (int)swSelectType_e.swSelFACES, "new_drw", "new_drawing_from_part", "", "", "");
                   swApp.AddMenuPopupItem2((int)swDocumentTypes_e.swDocPART, addinCookieID,	(int)swSelectType_e.swSelFACES,"open_dwg", "open_dwg", "","","");
                 swApp.AddMenuPopupItem2((int)swDocumentTypes_e.swDocASSEMBLY, addinCookieID,	(int)swSelectType_e.swSelFACES,"open_dwg", "open_dwg", "","","");
                       
                 // 添加复制模型DWG到剪贴板的菜单项
                 swApp.AddMenuPopupItem2((int)swDocumentTypes_e.swDocPART, addinCookieID, (int)swSelectType_e.swSelFACES, "copy_dwg_clipboard", "copy_model_dwg_to_clipboard", "", "", "");
                 swApp.AddMenuPopupItem2((int)swDocumentTypes_e.swDocASSEMBLY, addinCookieID, (int)swSelectType_e.swSelFACES, "copy_dwg_clipboard", "copy_model_dwg_to_clipboard", "", "", "");
                      
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[实体右键菜单] 初始化失败：{ex.Message}");
            }
        }

      

   

       
    }
}