namespace cad_plugin;

using System;
using System.Reflection;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

public class Class1
{
    [ComRegisterFunction]
    public static void RegisterFunction(Type t)
    {
        try
        {
            // 动态获取当前 AutoCAD 版本信息
            string acadVersion = Application.Version.ToString();
            
            Microsoft.Win32.RegistryKey localMachine = Microsoft.Win32.Registry.LocalMachine;
            // 遍历注册表找到正确的产品代码
            string baseKey = $"SOFTWARE\\Autodesk\\AutoCAD\\{acadVersion}";
            Microsoft.Win32.RegistryKey versionKey = localMachine.OpenSubKey(baseKey);
            
            if (versionKey != null)
            {
                foreach (string subKeyName in versionKey.GetSubKeyNames())
                {
                    // 通常产品代码是类似 ACAD-xxxx:xxx 的格式
                    if (subKeyName.StartsWith("ACAD-"))
                    {
                        string registryPath = $"{baseKey}\\{subKeyName}\\Applications\\mycad";
                        Microsoft.Win32.RegistryKey myProgram = localMachine.CreateSubKey(registryPath, true);
                        myProgram.SetValue("DESCRIPTION", "cad插件", Microsoft.Win32.RegistryValueKind.String);
                        myProgram.SetValue("LOADCTRLS", 2, Microsoft.Win32.RegistryValueKind.DWord);
                        myProgram.SetValue("LOADER", Assembly.GetExecutingAssembly().Location, Microsoft.Win32.RegistryValueKind.String);
                        myProgram.SetValue("MANAGED", 1, Microsoft.Win32.RegistryValueKind.DWord);
                        break; // 找到第一个匹配的就退出
                    }
                }
            }
        }
        catch (NullReferenceException nl)
        {
            Console.WriteLine("There was a problem registering this dll: SWattr is null. \n\"" + nl.Message + "\"");
            Console.WriteLine("There was a problem registering this dll: SWattr is null.\n\"" + nl.Message + "\"");
        }
        catch (System.Exception e)
        {
            Console.WriteLine(e.Message);
            Console.WriteLine("There was a problem registering the function: \n\"" + e.Message + "\"");
        }
    }

    [ComUnregisterFunction]
    public static void UnregisterFunction(Type t)
    {
        try
        {
            // 动态获取当前 AutoCAD 版本信息
            string acadVersion = Application.Version.ToString();
            
            Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
            // 遍历注册表找到正确的产品代码
            string baseKey = $"SOFTWARE\\Autodesk\\AutoCAD\\{acadVersion}";
            Microsoft.Win32.RegistryKey versionKey = hklm.OpenSubKey(baseKey);
            
            if (versionKey != null)
            {
                foreach (string subKeyName in versionKey.GetSubKeyNames())
                {
                    // 通常产品代码是类似 ACAD-xxxx:xxx 的格式
                    if (subKeyName.StartsWith("ACAD-"))
                    {
                        string keyname = $"{baseKey}\\{subKeyName}\\Applications\\mycad";
                        hklm.DeleteSubKey(keyname, false); // false 表示如果子键不存在不抛出异常
                        break; // 找到第一个匹配的就退出
                    }
                }
            }
        }
        catch (NullReferenceException nl)
        {
            Console.WriteLine("There was a problem unregistering this dll: " + nl.Message);
            Console.WriteLine("There was a problem unregistering this dll: \n\"" + nl.Message + "\"");
        }
        catch (System.Exception e)
        {
            Console.WriteLine("There was a problem unregistering this dll: " + e.Message);
            Console.WriteLine("There was a problem unregistering this dll: \n\"" + e.Message + "\"");
        }
    }
}