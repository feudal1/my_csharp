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
    


       
    [Command(1006, "打开 DWG工程图caxa", "打开 DWG工程图", "opendwgcaxa", (int)swDocumentTypes_e.swDocDRAWING)]
    private void Opencaxa()
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
            opendwg.run(swModel, swApp, true);


           
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"工程图dwg打开失败：{ex.Message}");
            swApp?.SendMsgToUser($"工程图dwg打开失败：{ex.Message}");
        }
    }
    [Command(1012, "打开 DWG工程图cad", "打开 DWG工程图", "opendwgcad", (int)swDocumentTypes_e.swDocDRAWING)]
    private void Opencad()
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
            opendwg.run(swModel, swApp, false);


           
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"工程图dwg打开失败：{ex.Message}");
            swApp?.SendMsgToUser($"工程图dwg打开失败：{ex.Message}");
        }
    }
    [Command(1009, "新建装配体工程图", "为当前装配体创建工程图并添加视图和 BOM 表", "asmnewdrw", (int)swDocumentTypes_e.swDocASSEMBLY, ShowOutputWindow = true)]
    private void AsmNewDrw()
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
                swApp.SendMsgToUser("请先打开一个装配体文档");
                return;
            }
           
            // 创建新装配体工程图
            Asm_new_drw.run(swApp, swModel);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"装配体工程图创建失败：{ex.Message}");
            swApp?.SendMsgToUser($"装配体工程图创建失败：{ex.Message}");
        }
    }
    [Command(1007, "装配体批量检查展开", "检查装配体中所有零件的展开情况", "asm2check", (int)swDocumentTypes_e.swDocASSEMBLY, ShowOutputWindow = true)]
    private void Asm2check()
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
                swApp.SendMsgToUser("请先打开一个装配体文档");
                return;
            }
           
            asm2do.run(swApp, swModel, (model, app) => {
                checkk_factor.run(app, model);
                return 0;
            });
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"装配体检查展开失败：{ex.Message}");
            swApp?.SendMsgToUser($"装配体检查展开失败：{ex.Message}");
        }
    }

    [Command(1010, "检查K因子", "检查当前零件的K因子设置是否正确", "checkkfactor", (int)swDocumentTypes_e.swDocPART, ShowOutputWindow = true)]
    private void CheckKFactor()
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
            
            checkk_factor.run(swApp, swModel);
            
            Debug.WriteLine("K因子检查完成");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"K因子检查失败：{ex.Message}");
            swApp?.SendMsgToUser($"K因子检查失败：{ex.Message}");
        }
    }

     [Command(1008, "装配体导出 BOM", "生成装配体 BOM 表并导出为 Excel", "asm2bom-pipe", (int)swDocumentTypes_e.swDocASSEMBLY, ShowOutputWindow = true)]
    private async void Asm2bom()
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
                swApp.SendMsgToUser("请先打开一个装配体文档");
                return;
            }
           
            await asm2bom.run(swApp, swModel, false);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"装配体 BOM 导出失败：{ex.Message}");
            swApp?.SendMsgToUser($"装配体 BOM 导出失败：{ex.Message}");
        }
    }
         [Command(1014, "装配体导出 BOM", "生成装配体 BOM 表并导出为 Excel", "asm2bom-sheet", (int)swDocumentTypes_e.swDocASSEMBLY, ShowOutputWindow = true)]
    private async void Asm2bom_sheet()
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
                swApp.SendMsgToUser("请先打开一个装配体文档");
                return;
            }
           
            await asm2bom.run(swApp, swModel, true);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"装配体 BOM 导出失败：{ex.Message}");
            swApp?.SendMsgToUser($"装配体 BOM 导出失败：{ex.Message}");
        }
    }

    [Command(1011, "批量导出 STEP", "批量导出装配体中所有零件为 STEP 格式", "asm2step", (int)swDocumentTypes_e.swDocASSEMBLY, ShowOutputWindow = true)]
    private void AsmBatchStep()
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
                swApp.SendMsgToUser("请先打开一个装配体文档");
                return;
            }
           
            asm2do.run(swApp, swModel, (model, app) =>
            {
                return one2step.run(model);
            });
            
            Debug.WriteLine("装配体零件 STEP 导出完成");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"装配体 STEP 导出失败：{ex.Message}");
            swApp?.SendMsgToUser($"装配体 STEP 导出失败：{ex.Message}");
        }
    }
    [Command(1013, "复制DWG文件到剪贴板", "将当前文档对应的DWG文件路径复制到剪贴板", "copyfile", 0, ShowOutputWindow = true)]
    private void CopyFileToClipboard()
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
                swApp.SendMsgToUser("请先打开一个文档");
                return;
            }

            string filePath = swModel.GetPathName();
            
            if (string.IsNullOrEmpty(filePath))
            {
                swApp.SendMsgToUser("当前文档尚未保存，请先保存文档");
                return;
            }

            // 如果是DRW文件，需要查找对应的DWG文件
            if (filePath.ToLower().EndsWith("drw"))
            {
                Console.WriteLine($"检测到DRW文件: {filePath}");
                // 使用与 drw2dwg 相同的路径计算逻辑
                string? directory = Path.GetDirectoryName(filePath);
                if (string.IsNullOrEmpty(directory))
                {
                    Console.WriteLine("无法获取文件目录");
                    swApp.SendMsgToUser("无法获取文件目录");
                    return;
                }

                string fileName = Path.GetFileNameWithoutExtension(filePath);
                Console.WriteLine($"文件名: {fileName}");
                string dwgPath = null;

                // 尝试多个可能的DWG路径
                // 1. 焊接图路径
                string weldingPath = Path.Combine(directory, "出图", "焊接图", fileName + ".dwg");
                Console.WriteLine($"检查焊接图路径: {weldingPath}, 存在: {File.Exists(weldingPath)}");
                if (File.Exists(weldingPath))
                {
                    dwgPath = weldingPath;
                    Console.WriteLine($"找到焊接图DWG: {dwgPath}");
                }
                else
                {
                    // 2. CNC路径
                    string cncPath = Path.Combine(directory, "出图", "CNC", fileName + ".dwg");
                    Console.WriteLine($"检查CNC路径: {cncPath}, 存在: {File.Exists(cncPath)}");
                    if (File.Exists(cncPath))
                    {
                        dwgPath = cncPath;
                        Console.WriteLine($"找到CNC DWG: {dwgPath}");
                    }
                    else
                    {
                        // 3. 工程图路径（可能有材质子目录）
                        string engDir = Path.Combine(directory, "出图", "工程图");
                        Console.WriteLine($"检查工程图目录: {engDir}, 存在: {Directory.Exists(engDir)}");
                        if (Directory.Exists(engDir))
                        {
                            // 搜索所有子目录中的DWG文件
                            var dwgFiles = Directory.GetFiles(engDir, fileName + ".dwg", SearchOption.AllDirectories);
                            Console.WriteLine($"在工程图目录中找到 {dwgFiles.Length} 个DWG文件");
                            if (dwgFiles.Length > 0)
                            {
                                dwgPath = dwgFiles[0];
                                Console.WriteLine($"找到工程图DWG: {dwgPath}");
                            }
                        }
                    }
                }

                if (string.IsNullOrEmpty(dwgPath))
                {
                    Console.WriteLine($"未找到对应的DWG文件");
                    Console.WriteLine("请先使用'工程图转DWG'命令转换");
                    swApp.SendMsgToUser("未找到对应的DWG文件，请先使用'工程图转DWG'命令转换");
                    return;
                }

                filePath = dwgPath;
                Console.WriteLine($"最终DWG路径: {filePath}");
            }
            if (filePath.ToLower().EndsWith(".sldprt") || filePath.ToLower().EndsWith(".prt"))
            {
                string partname = Path.GetFileNameWithoutExtension(filePath);
                string? currentDirectory = Path.GetDirectoryName(filePath);
                filePath = Path.Combine(currentDirectory, "step", $"{partname}.STEP");
            }
   

            if (!File.Exists(filePath))
            {
                swApp.SendMsgToUser($"DWG文件不存在: {filePath}");
                return;
            }

            System.Collections.Specialized.StringCollection fileList = new System.Collections.Specialized.StringCollection();
            fileList.Add(filePath);
            
            System.Windows.Forms.Clipboard.SetFileDropList(fileList);
            
            swApp.SendMsgToUser($"已将DWG文件复制到剪贴板:\n{Path.GetFileName(filePath)}");
            Debug.WriteLine($"DWG文件已复制到剪贴板: {filePath}");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"复制文件到剪贴板失败: {ex.Message}");
            swApp?.SendMsgToUser($"复制文件到剪贴板失败: {ex.Message}");
        }
    }

    [Command(1015, "装配体排版", "将装配体出图文件夹下的DWG文件进行自动排版", "asmdivider", (int)swDocumentTypes_e.swDocASSEMBLY, ShowOutputWindow = true)]
    private void AsmDivider()
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
                swApp.SendMsgToUser("请先打开一个装配体文档");
                return;
            }

            string assemblyPath = swModel.GetPathName();
            
            if (string.IsNullOrEmpty(assemblyPath))
            {
                swApp.SendMsgToUser("当前装配体尚未保存，请先保存文档");
                return;
            }

            string assemblyDir = Path.GetDirectoryName(assemblyPath);
            if (string.IsNullOrEmpty(assemblyDir))
            {
                swApp.SendMsgToUser("无法获取装配体目录");
                return;
            }

            string outputFolder = Path.Combine(assemblyDir, "出图");
            
            if (!Directory.Exists(outputFolder))
            {
                swApp.SendMsgToUser($"出图文件夹不存在: {outputFolder}");
                return;
            }

            draw_divider.process_subfolders_with_divider(outputFolder);
            
            Debug.WriteLine("装配体排版完成");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"装配体排版失败：{ex.Message}");
            swApp?.SendMsgToUser($"装配体排版失败：{ex.Message}");
        }
    }



    [Command(1016, "标折弯尺寸", "自动标注选中视图的折弯尺寸", "dimension_bends", (int)swDocumentTypes_e.swDocDRAWING, ShowOutputWindow = true)]
    private void DimensionBends()
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

            // 调用 benddim 标折弯尺寸
            benddim.AddBendDimensions(swApp);
            
            Debug.WriteLine("折弯尺寸标注完成");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"AddBendDimensions failed: {ex.Message}");
            swApp?.SendMsgToUser($"AddBendDimensions failed: {ex.Message}");
        }
    }




    
}}