namespace cad_plugin;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using cad_tools;

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
      
}
