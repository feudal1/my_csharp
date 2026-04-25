namespace cad_plugin;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using cad_tools;
using System.IO;
using System.Runtime.InteropServices;
using System;

public partial class CadPluginCommands
{
    [CommandMethod("HELLO")]
    public void HelloCommand()
    {
        Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
        ed.WriteMessage("\nHello from CAD Plugin!\n");
    }

    [CommandMethod("mergedwg")]
    public void DrawDividerCommand()
    {
        draw_divider.process_subfolders_with_divider();
    }
    
    [CommandMethod("COPYFILE")]
    public void CopyCurrentFileToClipboard()
    {
        try
        {
            // 获取当前活动的文档
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                Editor ed = Application.DocumentManager.MdiActiveDocument?.Editor;
                ed?.WriteMessage("\n错误：没有活动的文档！\n");
                return;
            }

            Editor editor = doc.Editor;
            
            // 先保存当前文档
            editor.WriteMessage("\n正在保存当前文档...\n");
            doc.Database.SaveAs(doc.Name, DwgVersion.Current);
            editor.WriteMessage("文档保存成功！\n");

            // 获取当前文件的路径
            string filePath = doc.Name;
            
            // 检查文件是否存在
            if (!File.Exists(filePath))
            {
                editor.WriteMessage($"\n错误：文件不存在: {filePath}\n");
                return;
            }

            // 使用Windows API将文件路径复制到剪贴板
            bool success = CopyFileToClipboard(filePath);
            
            // 显示结果消息
            if (success)
            {
                editor.WriteMessage($"\n已将文件复制到剪贴板: {Path.GetFileName(filePath)}\n");
                editor.WriteMessage($"完整路径: {filePath}\n");
            }
            else
            {
                editor.WriteMessage("\n复制到剪贴板失败！\n");
            }
        }
        catch (System.Exception ex)
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument?.Editor;
            ed?.WriteMessage($"\n复制文件到剪贴板时出错: {ex.Message}\n");
        }
    }
    
    /// <summary>
    /// 将文件路径复制到Windows剪贴板
    /// </summary>
    /// <param name="filePath">要复制的文件路径</param>
    /// <returns>是否成功复制</returns>
    private bool CopyFileToClipboard(string filePath)
    {
        try
        {
            // 使用COM接口来设置剪贴板数据
            // 这种方法适用于需要复制文件到剪贴板的场景
            System.Collections.Specialized.StringCollection fileList = new System.Collections.Specialized.StringCollection();
            fileList.Add(filePath);
            
            // 使用Windows Forms的Clipboard类
            System.Windows.Forms.Clipboard.SetFileDropList(fileList);
            
            return true;
        }
        catch (System.Exception ex)
        {
            // 如果Windows Forms方法失败，尝试其他方法
            System.Diagnostics.Debug.WriteLine($"设置剪贴板失败: {ex.Message}");
            return false;
        }
    }
      
}