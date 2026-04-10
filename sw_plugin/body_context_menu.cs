using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorksTools;
using System;
using System.Diagnostics;
using SolidWorks.Interop.swconst;
using System.Collections.Generic;
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
                      
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[实体右键菜单] 初始化失败：{ex.Message}");
            }
        }

      

   

       
    }
}