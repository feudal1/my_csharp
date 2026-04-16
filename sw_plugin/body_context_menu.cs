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

      public string export_selected_to_step()
      {
          ShowOutputWindow();
          try
          {
              Debug.WriteLine("开始导出STEP");
              if (swApp == null)
              {
                  Debug.WriteLine("SolidWorks 未初始化");
                  return "SolidWorks 未初始化";
              }

              var acswModel = (ModelDoc2)swApp.ActiveDoc;
              if (acswModel == null)
              {
                  Debug.WriteLine("没有打开的文档");
                  swApp.SendMsgToUser("请先打开一个文档");
                  return "没有打开的文档";
              }

              var (swModel, body) = get_select_body(acswModel);
              if (swModel == null || body == null)
              {
                  Debug.WriteLine("未选中实体");
                  swApp.SendMsgToUser("请先选中一个实体");
                  return "未选中实体";
              }

              string fullPath = swModel.GetPathName();
              if (string.IsNullOrEmpty(fullPath))
              {
                  Debug.WriteLine("文档尚未保存");
                  swApp.SendMsgToUser("请先保存文档");
                  return "文档尚未保存";
              }

              string partname = Path.GetFileNameWithoutExtension(fullPath);
              string? currentDirectory = Path.GetDirectoryName(fullPath);
              
              string outputPath = Path.Combine(currentDirectory, "step", $"{partname}.STEP");
              string outputDirectory = Path.GetDirectoryName(outputPath);
              
              if (!Directory.Exists(outputDirectory))
              {
                  Directory.CreateDirectory(outputDirectory);
              }

              var result = swModel.SaveAs3(outputPath, 0, 2);
              
              Debug.WriteLine($"STEP导出结果: {result}, 路径: {outputPath}");
              swApp.SendMsgToUser($"STEP文件已导出到: {outputPath}");
              return $"STEP文件已导出: {outputPath}";
          }
          catch (Exception ex)
          {
              Debug.WriteLine($"导出STEP失败：{ex.Message}");
              swApp?.SendMsgToUser($"导出STEP失败：{ex.Message}");
              return $"导出STEP失败：{ex.Message}";
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

                
         
  
                            swApp.AddMenuPopupItem2((int)swDocumentTypes_e.swDocPART, addinCookieID, (int)swSelectType_e.swSelFACES, "new_drw", "new_drawing_from_part", "", "", "");
                swApp.AddMenuPopupItem2((int)swDocumentTypes_e.swDocASSEMBLY, addinCookieID, (int)swSelectType_e.swSelFACES, "new_drw", "new_drawing_from_part", "", "", "");

                swApp.AddMenuPopupItem2((int)swDocumentTypes_e.swDocPART, addinCookieID, (int)swSelectType_e.swSelFACES, "export_step", "export_selected_to_step", "", "", "");
                swApp.AddMenuPopupItem2((int)swDocumentTypes_e.swDocASSEMBLY, addinCookieID, (int)swSelectType_e.swSelFACES, "export_step", "export_selected_to_step", "", "", "");

                      
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[实体右键菜单] 初始化失败：{ex.Message}");
            }
        }

      

   

       
    }
}