using System;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace tools
{
     class Connect
    {
        static public SldWorks? run()
        { 
            var progId = "SldWorks.Application";

            if (!OperatingSystem.IsWindows())
            {
                Console.WriteLine("错误：此功能仅支持 Windows 平台。");
                return null;
            }

            try
            {
                // 尝试获取正在运行的 SolidWorks 实例
                object comObject = com_tools.ComHelper.GetActiveObject(progId);
                return (SldWorks)comObject;
            }
            catch (COMException)
            {
                // 如果没有运行中的实例，创建新实例
                Type? comType = Type.GetTypeFromProgID(progId);
                if (comType == null)
                {
                    Console.WriteLine("错误：无法获取 SolidWorks 应用程序类型。");
                    return null;
                }

                object? newComObject = Activator.CreateInstance(comType);
                if (newComObject == null)
                {
                    Console.WriteLine("错误：无法创建 SolidWorks 应用程序实例。");
                    return null;
                }

                return (SldWorks)newComObject;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"连接 SolidWorks 失败：{ex.Message}");
                return null;
            }
        }

       
    }
    }
