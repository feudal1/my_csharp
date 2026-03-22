using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace tools
{
    public class asm2do
    {
        static private Dictionary<string, int> docnameCount = new Dictionary<string, int>();
        static private readonly object _swApiLock = new object();

        public delegate void SolidWorksAction( ModelDoc2 swModel,SldWorks swApp=null);

        static public string[]? run(SldWorks swApp, ModelDoc2 swModel, SolidWorksAction action)
        {
            try
            {

                if (swModel == null)
                {
                    Console.WriteLine("错误：没有打开任何文档。");
                    return null;
                }

                // 确保是装配体
                if (swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    Console.WriteLine("错误：当前文档不是装配体。");
                    return null;
                }

                AssemblyDoc swAssembly = (AssemblyDoc)swModel;

                // 获取顶层组件
                object[] topComponents = (object[])swAssembly.GetComponents(false);

                Console.WriteLine($"正在扫描装配体... 共 {topComponents.Length} 个顶层组件");

                // 单线程顺序处理
                foreach (object compObj in topComponents)
                {
                    Component2 topComp = (Component2)compObj;
                    TraverseComponent(swApp, topComp, action);
                }



                foreach (var kvp in docnameCount.OrderBy(x => x.Key))
                {
                    Console.WriteLine($"{kvp.Key}: {kvp.Value} 次");
                }

                List<string> result = new List<string>(docnameCount.Keys);
                return result.ToArray();

            }
            catch (Exception ex)
            {
                Console.WriteLine($"发生错误: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }

            return null;
        }

// 递归遍历组件
        static private void TraverseComponent(SldWorks swApp, Component2 component, SolidWorksAction action)
        {
            // 获取子组件
            object[] children = (object[])component.GetChildren();

            if (children != null && children.Length > 0)
            {
                foreach (object childObj in children)
                {
                    Component2 childComp = (Component2)childObj;
                    TraverseComponent(swApp, childComp, action);
                }
            }
            else
            {
                ModelDoc2 doc = (ModelDoc2)component.GetModelDoc2();
                if (doc != null)
                {
                    doc.Visible = true;
                    string? pathName = doc.GetPathName();
                    if (string.IsNullOrEmpty(pathName))
                    {
                        Console.WriteLine($"警告：组件 {component.Name2} 未保存或无法获取路径。");
                        // 即使没有路径，也要尝试关闭文档以释放资源
                        swApp.CloseDoc("");
                        return;
                    }

                    string docname = pathName.Replace(@"\\", "/");

                    // 添加到字典
                    if (!docnameCount.ContainsKey(docname))
                    {
                        docnameCount[docname] = 0;
                    }

                    docnameCount[docname]++;

                    // 只处理第一次遇到的零件
                    if (docnameCount[docname] == 1)
                    {

                        Debug.WriteLine($"正在处理 {docname}...");
                        action.Invoke( doc,swApp);




                    }

                    // 成功处理完后，关闭文档
                    swApp.CloseDoc(pathName);



                }
            }
        }
    }
}

