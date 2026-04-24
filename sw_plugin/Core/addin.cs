using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorksTools;
using System;
using System.IO;
using System.Runtime.InteropServices;
using tools;
using System.Diagnostics;
using SolidWorks.Interop.swconst;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace SolidWorksAddinStudy
{
    [Guid("D9C5D3A4-3B9F-4ACF-BC19-6D77D39C47CD"), ComVisible(true)]
    [SwAddin(
    Description = "SolidWorksAddinStudy description",
    Title = "SolidWorksAddinStudy",
    LoadAtStartup = true
    )]
    public partial class AddinStudy : ISwAddin
    {
        
        private static SldWorks? swApp;
        private static ICommandManager? iCmdMgr;
        private static int addinCookieID;
        private static bool consoleOpened = false;
        private static ConsoleOutputForm? consoleForm;
        private static System.Threading.SynchronizationContext? mainSynchronizationContext;
        
        // 任务窗格相关
        private static ITaskpaneView? pTaskPanView;
        private static PartStatusControl? TaskPanWinFormControl;
        private static ITaskpaneView? workProjectTaskPaneView;
        private static WorkProjectTaskPaneControl? workProjectTaskPaneControl;
        private static ITaskpaneView? equationModelTaskPaneView;
        private static EquationModelTaskPaneControl? equationModelTaskPaneControl;


        /// <summary>
        /// 显示输出窗口
        /// </summary>
        public static void ShowOutputWindow()
        {
            if (consoleForm == null || consoleForm.IsDisposed)
            {
                consoleForm = new ConsoleOutputForm();
               // 设置窗口置顶
                consoleForm.Show();
                consoleForm.StartIntercept();
            }
            else
            {
                if (consoleForm.WindowState == FormWindowState.Minimized)
                {
                    consoleForm.WindowState = FormWindowState.Normal;
                }
              
                consoleForm.BringToFront();
            }
        }

        /// <summary>
        /// 关闭输出窗口
        /// </summary>
        public static void CloseOutputWindow()
        {
            if (consoleForm != null && !consoleForm.IsDisposed)
            {
                consoleForm.StopIntercept();
                consoleForm.Close();
                consoleForm = null;
            }
        }

        [DllImport("kernel32.dll")]
        private static extern bool AllocConsole();

        [DllImport("kernel32.dll")]
        private static extern bool FreeConsole();
        
        [DllImport("kernel32.dll")]
        private static extern IntPtr GetConsoleWindow();
        
        [DllImport("user32.dll")]
        private static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);
        
        [DllImport("user32.dll")]
        private static extern bool RemoveMenu(IntPtr hMenu, uint uPosition, uint uFlags);
        
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        
        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);
        
        private const int SW_HIDE = 0;
        private const int SW_SHOW = 5;
        private const uint SC_CLOSE = 0xF060;
        private const uint MF_BYCOMMAND = 0x00000000;

        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            swApp = (SldWorks)ThisSW;
            mainSynchronizationContext = System.Threading.SynchronizationContext.Current;
            addinCookieID = Cookie;
            swApp.SetAddinCallbackInfo(0, this, addinCookieID);
            iCmdMgr = swApp.GetCommandManager(addinCookieID);
            
            Debug.WriteLine("插件已加载...");
            
            // 初始化全局 SolidWorks 上下文
            var swModel = (ModelDoc2)swApp.ActiveDoc;
   
            
            // 初始化命令注册表
            InitializeCommandRegistry();
            

 
            AddCommandMgr();
            ShowWelcomeImage();
            
            PopupMenuInitialize();
            
            // 初始化任务窗格
            InitializeTaskPane();

            // 设置asm2bom的任务窗格更新回调
            tools.asm2bom.TaskPaneUpdateCallback = OnBomDataUpdated;

            StartCommandBridgeServer();

            return true;
        }
/// <summary>
        /// 获取程序集构建时间
        /// </summary>
        private DateTime GetBuildTime()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            string filePath = assembly.Location;
            return new FileInfo(filePath).LastWriteTime;
        }
        
        private void ShowWelcomeImage()
        {
            try
        {
              string pluginDir = Path.GetDirectoryName(typeof(AddinStudy).Assembly.Location);
            string imagePath = Path.Combine(pluginDir, "welcome.png");
                if (!File.Exists(imagePath))
            {
                    System.Windows.Forms.MessageBox.Show($"图片文件不存在：{imagePath}");
                return;
                }
              

                using (Image img = Image.FromFile(imagePath))
            {
                // 获取构建时间
                DateTime buildTime = GetBuildTime();
                string versionText = $"版本: {buildTime:yyyy-MM-dd HH:mm}";
                
                Form form = new Form
                {
                    Text = versionText,
                    Size = new Size(img.Width + 40, img.Height + 90),
                    StartPosition = FormStartPosition.CenterScreen,
                    TopMost = true,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    MaximizeBox = false,
                    MinimizeBox = false,
               
                };
                
                PictureBox pictureBox = new PictureBox
                {
                    Image = img,
                    SizeMode = PictureBoxSizeMode.AutoSize,
                    Location = new Point(10, 0)
                };
                
             
          
                
                Button okButton = new Button
                {
                    Text = "清空所有工程文件？ (5)",
                    DialogResult = DialogResult.OK,
                    Location = new Point((form.Width - 200) / 2, form.Height - 80),
                    Size = new Size(200, 30)
                };
                
                form.Controls.Add(pictureBox);
             
                form.Controls.Add(okButton);
                form.AcceptButton = okButton;
                      Timer timer = new Timer
                {
                    Interval = 1000
                };
                int countdown = 5;
                timer.Tick += (s, e) =>
                {
                    countdown--;
                    if (countdown <= 0)
                    {
                        timer.Stop();
                        form.Close();
                    }
                    else
                    {
                        okButton.Text = $"清空所有工程文件？  ({countdown})";
                    }
                };
                timer.Start();
                form.ShowDialog();
            }
        }
        catch (Exception ex)
        {
                System.Windows.Forms.MessageBox.Show($"显示图片失败：{ex.Message}");
        }
    }
        public bool DisconnectFromSW()
        {
            Debug.WriteLine("插件已卸载");
            StopCommandBridgeServer();

            // 清理任务窗格
            CleanupTaskPane();
        
        
            return true;
        }


        /// <summary>
        /// 决定此命令在该环境下是否可用
        /// </summary>
        public int EnableFunction(string data)
        {
            return 1;
        }

       


  

        /// <summary>
        /// 比较 ID 集
        /// </summary>
        private bool CompareIDs(int[] storedIDs, int[] addinIDs)
        {
            List<int> storedList = new List<int>(storedIDs);
            List<int> addinList = new List<int>(addinIDs);

            addinList.Sort();
            storedList.Sort();

            if (addinList.Count != storedList.Count)
            {
                return false;
            }

            for (int i = 0; i < addinList.Count; i++)
            {
                if (addinList[i] != storedList[i])
                {
                    return false;
                }
            }
            return true;
        }

 #region SolidWorks Registration
 
        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {
            #region Get Custom Attribute: SwAddinAttribute
 
            SwAddinAttribute SWattr = null;
            Type type = typeof(AddinStudy );
 
            foreach (System.Attribute attr in type.GetCustomAttributes(false))
            {
                if (attr is SwAddinAttribute)
                {
                    SWattr = attr as SwAddinAttribute;
                    break;
                }
            }
 
            #endregion Get Custom Attribute: SwAddinAttribute
 
            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;
 
                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                Microsoft.Win32.RegistryKey addinkey = hklm.CreateSubKey(keyname);
                addinkey.SetValue(null, 0);
 
                addinkey.SetValue("Description", SWattr.Description);
                addinkey.SetValue("Title", SWattr.Title);
 
                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                addinkey = hkcu.CreateSubKey(keyname);
                addinkey.SetValue(null, Convert.ToInt32(SWattr.LoadAtStartup), Microsoft.Win32.RegistryValueKind.DWord);
            }
            catch (System.NullReferenceException nl)
            {
                Debug.WriteLine("There was a problem registering this dll: SWattr is null. \n\"" + nl.Message + "\"");
                
            }
            catch (System.Exception e)
            {
                Debug.WriteLine(e.Message);
 
            }
        }
 
        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;
 
                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                hklm.DeleteSubKey(keyname);
 
                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                hkcu.DeleteSubKey(keyname);
            }
            catch (System.NullReferenceException nl)
            {
                Debug.WriteLine("There was a problem unregistering this dll: " + nl.Message);
                
            }
            catch (System.Exception e)
            {
                Debug.WriteLine("There was a problem unregistering this dll: " + e.Message);
                
            }
        }
 
        #endregion SolidWorks Registration

        /// <summary>
        /// 初始化任务窗格
        /// </summary>
        private void InitializeTaskPane()
        {
            try
            {
                if (swApp == null) return;
                
                // 如果已经存在任务窗格，先清理
                CleanupTaskPane();
                
                // 创建零件状态任务窗格视图
                pTaskPanView = swApp.CreateTaskpaneView2("", "零件处理状态");
                
                if (pTaskPanView != null)
                {
                    TaskPanWinFormControl = new PartStatusControl(swApp);
                    pTaskPanView.DisplayWindowFromHandlex64(TaskPanWinFormControl.Handle.ToInt64());
                    Debug.WriteLine("零件状态任务窗格已初始化");
                }

                // 创建工作项目记录任务窗格
                workProjectTaskPaneView = swApp.CreateTaskpaneView2("", "工作项目记录");
                if (workProjectTaskPaneView != null)
                {
                    workProjectTaskPaneControl = new WorkProjectTaskPaneControl();
                    workProjectTaskPaneView.DisplayWindowFromHandlex64(workProjectTaskPaneControl.Handle.ToInt64());
                    workProjectTaskPaneControl.PromptFollowUpRemindersAtStartup();
                    Debug.WriteLine("工作项目记录任务窗格已初始化");
                }

                equationModelTaskPaneView = swApp.CreateTaskpaneView2("", "机型方程式");
                if (equationModelTaskPaneView != null)
                {
                    equationModelTaskPaneControl = new EquationModelTaskPaneControl();
                    equationModelTaskPaneView.DisplayWindowFromHandlex64(equationModelTaskPaneControl.Handle.ToInt64());
                    Debug.WriteLine("机型方程式任务窗格已初始化");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"初始化任务窗格失败: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 清理任务窗格
        /// </summary>
        private void CleanupTaskPane()
        {
            try
            {
                if (TaskPanWinFormControl != null)
                {
                    TaskPanWinFormControl.Dispose();
                    TaskPanWinFormControl = null;
                }

                if (workProjectTaskPaneControl != null)
                {
                    workProjectTaskPaneControl.Dispose();
                    workProjectTaskPaneControl = null;
                }

                if (equationModelTaskPaneControl != null)
                {
                    equationModelTaskPaneControl.Dispose();
                    equationModelTaskPaneControl = null;
                }
                
                if (pTaskPanView != null)
                {
                    pTaskPanView.DeleteView();
                    pTaskPanView = null;
                }

                if (workProjectTaskPaneView != null)
                {
                    workProjectTaskPaneView.DeleteView();
                    workProjectTaskPaneView = null;
                }

                if (equationModelTaskPaneView != null)
                {
                    equationModelTaskPaneView.DeleteView();
                    equationModelTaskPaneView = null;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"清理任务窗格失败: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 获取任务窗格控件实例
        /// </summary>
        public static PartStatusControl? GetTaskPaneControl()
        {
            return TaskPanWinFormControl;
        }

        public static WorkProjectTaskPaneControl? GetWorkProjectTaskPaneControl()
        {
            return workProjectTaskPaneControl;
        }

        public static EquationModelTaskPaneControl? GetEquationModelTaskPaneControl()
        {
            return equationModelTaskPaneControl;
        }

        public static SldWorks? GetSwApp()
        {
            return swApp;
        }

        public static System.Threading.SynchronizationContext? GetMainSynchronizationContext()
        {
            return mainSynchronizationContext;
        }
        
        /// <summary>
        /// BOM数据更新回调处理
        /// </summary>
        private static void OnBomDataUpdated(List<object> bomDataList)
        {
            try
            {
                var taskPaneControl = GetTaskPaneControl();
                if (taskPaneControl == null || bomDataList == null || bomDataList.Count == 0) return;
                
                // 将匿名对象转换为BomPartInfo列表
                var partInfoList = new List<PartStatusInfo>();
                
                foreach (var item in bomDataList)
                {
                    // 使用反射获取属性值
                    var partName = item.GetType().GetProperty("PartName")?.GetValue(item)?.ToString() ?? "";
                    var partType = item.GetType().GetProperty("PartType")?.GetValue(item)?.ToString() ?? "";
                    var dimension = item.GetType().GetProperty("Dimension")?.GetValue(item)?.ToString() ?? "";
                    var isDrawn = item.GetType().GetProperty("IsDrawn")?.GetValue(item)?.ToString() ?? "未出图";
                    var quantity = item.GetType().GetProperty("Quantity")?.GetValue(item)?.ToString() ?? "";
                    
                    partInfoList.Add(new PartStatusInfo
                    {
                        PartName = partName,
                        PartType = partType,
                        Dimension = dimension,
                        IsDrawn = isDrawn,
                        Quantity = quantity
                    });
                }
                
                // 在UI线程上更新
                if (taskPaneControl.InvokeRequired)
                {
                    taskPaneControl.Invoke(new Action(() => taskPaneControl.LoadFromBomData(partInfoList)));
                }
                else
                {
                    taskPaneControl.LoadFromBomData(partInfoList);
                }
                
                Debug.WriteLine($"任务窗格已更新，共 {partInfoList.Count} 个零件");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"处理BOM数据更新失败: {ex.Message}");
            }
        }


}}

