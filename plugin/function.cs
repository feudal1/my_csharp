using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorksTools;
using System;
using System.IO;
using System.Runtime.InteropServices;
using tools;
using cad_tools;
using System.Diagnostics;
using SolidWorks.Interop.swconst;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Reflection;
using System.Threading.Tasks;
using System.Text;

using System.Linq;

   namespace SolidWorksAddinStudy
{
   
    public partial class AddinStudy 
{
  
 


    [Command(1001, "导出展开", "装配体每个零件批量展开", "asm2export", (int)swDocumentTypes_e.swDocASSEMBLY, ShowOutputWindow = true)]
    private void Asm2export()
    {
        try
        {
            if (swApp == null)
            {
                Debug.WriteLine("SolidWorks 未初始化");
                return;
            }

            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            
            if (swModel == null)
            {
                Debug.WriteLine("没有打开的文档");
                swApp.SendMsgToUser("请先打开一个工程图文档");
                return;
            }
           
            asm2do.run(swApp, swModel, (model, app) => {
                checkk_factor.run(app,model);
                return exportdwg2_body.run(model);
            });


           
         
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"工程图转换失败：{ex.Message}");
            swApp?.SendMsgToUser($"工程图转换失败：{ex.Message}");
        }
    }
    [Command(1004, "导出展开", "导出展开", "exportdwg", (int)swDocumentTypes_e.swDocPART, ShowOutputWindow = true)]
    private void exportdwg()
    {
        try
        {
            if (swApp == null)
            {
                Debug.WriteLine("SolidWorks 未初始化");
                return;
            }

            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            
            if (swModel == null)
            {
                Debug.WriteLine("没有打开的文档");
                swApp.SendMsgToUser("请先打开一个工程图文档");
                return;
            }
            checkk_factor.run(swApp,swModel);
            exportdwg2_body.run(swModel);






        }
        catch (Exception ex)
        {
            Debug.WriteLine($"工程图转换失败：{ex.Message}");
            swApp?.SendMsgToUser($"工程图转换失败：{ex.Message}");
        }
    }
  
    [Command(1002, "工程图转 DWG", "将当前工程图转换为 DWG 格式并打开", "drw2dwg", (int)swDocumentTypes_e.swDocDRAWING)]
    private void Drw2Dwg()
    {
        try
        {
            if (swApp == null)
            {
                Debug.WriteLine("SolidWorks 未初始化");
                return;
            }

            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            
            if (swModel == null)
            {
                Debug.WriteLine("没有打开的文档");
                swApp.SendMsgToUser("请先打开一个工程图文档");
                return;
            }
           
            // 使用 share 项目中的 drw2dwg 方法转换 DWG
            var dwgFileName = drw2dwg.run(swModel, swApp);


           
            Debug.WriteLine($"工程图已转换为 DWG: {dwgFileName}");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"工程图转换失败：{ex.Message}");
            swApp?.SendMsgToUser($"工程图转换失败：{ex.Message}");
        }
    }

    [Command(1003, "新建工程图", "为当前零件创建工程图并添加视图", "newdrw", (int)swDocumentTypes_e.swDocPART)]
    private void NewDrw()
    {
        try
        {
            if (swApp == null)
            {
                Debug.WriteLine("SolidWorks 未初始化");
                return;
            }

            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            if (swModel == null)
            {
                Debug.WriteLine("没有打开的文档");
                swApp.SendMsgToUser("请先打开一个零件文档");
                return;
            }

            // 添加名称到自定义信息
            add_name2info.run(swModel);
            
            // 创建新工程图
            New_drw.run(swApp, swModel);
            
            Debug.WriteLine("工程图已创建");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"创建工程图失败：{ex.Message}");
            swApp?.SendMsgToUser($"创建工程图失败：{ex.Message}");
        }
    }
    [Command(1005, "新建工程图", "为当前零件创建工程图并添加视图", "newdrw2", (int)swDocumentTypes_e.swDocPART)]
    private void NewDrw2()
    {
        try
        {
            if (swApp == null)
            {
                Debug.WriteLine("SolidWorks 未初始化");
                return;
            }

            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            if (swModel == null)
            {
                Debug.WriteLine("没有打开的文档");
                swApp.SendMsgToUser("请先打开一个零件文档");
                return;
            }

      
            
            // 创建新工程图
            New_drw2.run(swApp, swModel);
            
            Debug.WriteLine("工程图已创建");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"创建工程图失败：{ex.Message}");
            swApp?.SendMsgToUser($"创建工程图失败：{ex.Message}");
        }
    }
    

 

       
    [Command(1007, "打开 DWG工程图", "打开 DWG工程图", "opendwg", (int)swDocumentTypes_e.swDocDRAWING)]
    private void Openwg()
    {
        try
        {
            if (swApp == null)
            {
                Debug.WriteLine("SolidWorks 未初始化");
                return;
            }

            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            
            if (swModel == null)
            {
                Debug.WriteLine("没有打开的文档");
                swApp.SendMsgToUser("请先打开一个工程图文档");
                return;
            }
           
            // 使用 share 项目中的 drw2dwg 方法转换 DWG
            opendwg.run(swModel, swApp);


           
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"工程图dwg打开失败：{ex.Message}");
            swApp?.SendMsgToUser($"工程图dwg打开失败：{ex.Message}");
        }
    }
    [Command(1006, "工程图转 DWGr12", "将当前工程图转换为 DWG 格式并打开", "drw2dwgr12", (int)swDocumentTypes_e.swDocDRAWING)]
    private void Drw2Dwgr12()
    {
        try
        {
            if (swApp == null)
            {
                Debug.WriteLine("SolidWorks 未初始化");
                return;
            }

            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
            
            if (swModel == null)
            {
                Debug.WriteLine("没有打开的文档");
                swApp.SendMsgToUser("请先打开一个工程图文档");
                return;
            }
           
            // 使用 share 项目中的 drw2dwg 方法转换 DWG
            var dwgFileName = drw2dwgr12.run(swModel, swApp);


           
            Debug.WriteLine($"工程图已转换为 DWG: {dwgFileName}");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"工程图转换失败：{ex.Message}");
            swApp?.SendMsgToUser($"工程图转换失败：{ex.Message}");
        }
    }

     

}}