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

              public string ShowBodyInfo()
              {
                  var swModel = (ModelDoc2)swApp.ActiveDoc;
                  get_select_type.run(swModel);
            return "a";

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

                
                swApp.AddMenuPopupItem2((int)swDocumentTypes_e.swDocPART, addinCookieID,	(int)swSelectType_e.swSelFACES,"打开dwg", "ShowBodyInfo", "","打开dwg","a");
                 swApp.AddMenuPopupItem2((int)swDocumentTypes_e.swDocASSEMBLY, addinCookieID,	(int)swSelectType_e.swSelFACES,"打开dwg", "ShowBodyInfo", "","打开dwg","a");
                 
              
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[实体右键菜单] 初始化失败：{ex.Message}");
            }
        }

      

   

       
    }
}
