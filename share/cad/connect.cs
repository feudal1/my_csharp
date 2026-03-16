using System;
using System.IO;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using com_tools;

namespace cad_tools
{
    class CadConnect
    {
        static public AcadApplication? GetOrCreateInstance()
        { 
            var progId = "AutoCAD.Application";

       
            try
            {
                // 尝试获取正在运行的 AutoCAD 实例
                object comObject =  ComHelper.GetActiveObject(progId);
                return (AcadApplication)comObject;
            }
            catch (COMException)
            {
                // 如果没有运行中的实例，创建新实例
                Type? comType = Type.GetTypeFromProgID(progId);
                if (comType == null)
                {
                    Console.WriteLine("错误：无法获取 AutoCAD 应用程序类型。");
                    return null;
                }

                object? comObject = Activator.CreateInstance(comType);
                if (comObject == null)
                {
                    Console.WriteLine("错误：无法创建 AutoCAD 应用程序实例。");
                    return null;
                }

                return (AcadApplication)comObject;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"连接 AutoCAD 失败：{ex.Message}");
                return null;
            }
        }
    }
}