using System;
using System.Windows.Forms;
using System.Collections.Generic;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace SolidWorksAddinStudy
{
    /// <summary>
    /// 零件处理状态用户控件
    /// </summary>
    [ComVisible(true)]
    [Guid("36B3CA37-FEC7-4E85-9E9A-1A7561242B04")]
    [ProgId("SolidWorksAddinStudy.PartStatusControl")]
    public partial class PartStatusControl : UserControl
    {
        private DataGridView statusGrid;
        private Label infoLabel;
        private Button refreshButton;
        private Button compressNonSheetButton;
        
        // 存储零件状态数据
        private List<PartStatusInfo> partStatusList = new List<PartStatusInfo>();
        
        // SolidWorks应用实例
        private SldWorks swApp;
        
        public PartStatusControl()
        {
            this.swApp = AddinStudy.GetSwApp();
            InitializeComponent();
        }

        public PartStatusControl(SldWorks swApp)
        {
            this.swApp = swApp;
            InitializeComponent();
        }

        public void SetSwApp(SldWorks swApp)
        {
            this.swApp = swApp;
        }
        
        private void InitializeComponent()
        {
            this.Dock = DockStyle.Fill;
            this.MinimumSize = new System.Drawing.Size(400, 400);
            
            // 创建顶部面板
            Panel header = new Panel();
            header.Dock = DockStyle.Top;
            header.Height = 50;
            header.Padding = new Padding(10, 10, 10, 10);
            
            // 创建信息标签
            infoLabel = new Label();
            infoLabel.Text = "零件处理状态监控";
            infoLabel.Location = new System.Drawing.Point(0, 15);
            infoLabel.Size = new System.Drawing.Size(200, 20);
            infoLabel.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Bold);
            header.Controls.Add(infoLabel);
            
            // 创建刷新按钮
            refreshButton = new Button();
            refreshButton.Text = "刷新数据";
            refreshButton.Location = new System.Drawing.Point(480, 10);
            refreshButton.Size = new System.Drawing.Size(90, 30);
            refreshButton.Font = new System.Drawing.Font("Microsoft YaHei", 9F);
            refreshButton.Click += RefreshButton_Click;
            header.Controls.Add(refreshButton);

            // 创建压缩非钣金按钮
            compressNonSheetButton = new Button();
            compressNonSheetButton.Text = "压缩非钣金";
            compressNonSheetButton.Location = new System.Drawing.Point(580, 10);
            compressNonSheetButton.Size = new System.Drawing.Size(110, 30);
            compressNonSheetButton.Font = new System.Drawing.Font("Microsoft YaHei", 9F);
            compressNonSheetButton.Click += CompressNonSheetButton_Click;
            header.Controls.Add(compressNonSheetButton);
            
            // 创建DataGridView显示状态
            statusGrid = new DataGridView();
            statusGrid.Dock = DockStyle.Fill;
            statusGrid.AllowUserToAddRows = false;
            statusGrid.AllowUserToDeleteRows = false;
            statusGrid.ReadOnly = true;
            statusGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            statusGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            statusGrid.MultiSelect = false;
            
            // 显式创建列，避免索引表达式在某些宿主环境下解析失败
            statusGrid.Columns.Clear();
            statusGrid.Columns.Add(CreateTextColumn("零件名称", 150));
            statusGrid.Columns.Add(CreateTextColumn("零件类型", 80));
            statusGrid.Columns.Add(CreateTextColumn("规格尺寸", 120));
            statusGrid.Columns.Add(CreateTextColumn("是否出图", 80));
            statusGrid.Columns.Add(CreateTextColumn("数量", 60));
            
            // 添加单元格点击事件
            statusGrid.CellClick += StatusGrid_CellClick;
            
            // 添加控件（注意顺序：后添加的在上层）
            this.Controls.Add(statusGrid);
            this.Controls.Add(header);
        }
        
        /// <summary>
        /// 刷新按钮点击事件
        /// </summary>
        private void RefreshButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (swApp == null)
                {
                    MessageBox.Show("SolidWorks未连接", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                
                ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
                if (swModel == null)
                {
                    MessageBox.Show("没有打开的文档", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                if (swModel.GetType() != (int)SolidWorks.Interop.swconst.swDocumentTypes_e.swDocASSEMBLY)
                {
                    MessageBox.Show("当前文档不是装配体", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                Debug.WriteLine("开始手动刷新BOM数据...");
                infoLabel.Text = "正在刷新数据...";
                refreshButton.Enabled = false;
                
                // 异步调用asm2bom生成BOM
                System.Threading.Tasks.Task.Run(async () =>
                {
                    try
                    {
                        await tools.asm2bom.run(swApp, swModel, "零件", false);
                        Debug.WriteLine("BOM数据刷新完成");
                        
                        // 刷新完成后重新启用按钮
                        this.Invoke(new Action(() =>
                        {
                            refreshButton.Enabled = true;
                            infoLabel.Text = $"零件处理状态监控 (共 {partStatusList.Count} 条记录)";
                        }));
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"刷新BOM失败: {ex.Message}");
                        this.Invoke(new Action(() =>
                        {
                            infoLabel.Text = $"刷新失败: {ex.Message}";
                            refreshButton.Enabled = true;
                        }));
                    }
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"刷新按钮点击失败: {ex.Message}");
                MessageBox.Show($"刷新失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                refreshButton.Enabled = true;
            }
        }

        /// <summary>
        /// 压缩全部非钣金件
        /// </summary>
        private void CompressNonSheetButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (swApp == null)
                {
                    MessageBox.Show("SolidWorks未连接", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
                if (swModel == null)
                {
                    MessageBox.Show("没有打开的文档", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    MessageBox.Show("当前文档不是装配体", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (partStatusList.Count == 0)
                {
                    MessageBox.Show("任务窗格暂无零件数据，请先点击“刷新数据”", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                HashSet<string> nonSheetPartNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var part in partStatusList)
                {
                    if (!string.Equals(part.PartType, "钣金件", StringComparison.OrdinalIgnoreCase))
                    {
                        nonSheetPartNames.Add(part.PartName);
                    }
                }

                if (nonSheetPartNames.Count == 0)
                {
                    MessageBox.Show("当前任务窗格中没有可压缩的非钣金件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                AssemblyDoc swAssembly = (AssemblyDoc)swModel;
                object[] components = (object[])swAssembly.GetComponents(false);
                if (components == null || components.Length == 0)
                {
                    MessageBox.Show("当前装配体没有可处理的组件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                int targetCount = 0;
                int successCount = 0;
                int failCount = 0;

                foreach (object compObj in components)
                {
                    Component2 component = (Component2)compObj;
                    string normalizedName = NormalizeComponentName(component.Name2);
                    if (!nonSheetPartNames.Contains(normalizedName))
                    {
                        continue;
                    }

                    targetCount++;
                    try
                    {
                        int result = component.SetSuppression2((int)swComponentSuppressionState_e.swComponentFullyLightweight);
                        if (result == (int)swSuppressionError_e.swSuppressionChangeOk)
                        {
                            successCount++;
                        }
                        else
                        {
                            failCount++;
                            Debug.WriteLine($"压缩失败: {component.Name2}, 返回码={result}");
                        }
                    }
                    catch (Exception ex)
                    {
                        failCount++;
                        Debug.WriteLine($"压缩组件异常: {component.Name2}, {ex.Message}");
                    }
                }

                swModel.EditRebuild3();
                string message = $"压缩完成：目标 {targetCount} 个，成功 {successCount} 个，失败 {failCount} 个";
                infoLabel.Text = message;
                swApp.SendMsgToUser(message);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"压缩非钣金件失败: {ex.Message}");
                MessageBox.Show($"压缩失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 从命令刷新BOM数据（供外部调用）
        /// </summary>
        public void RefreshFromCommand()
        {
            try
            {
                if (swApp == null)
                {
                    MessageBox.Show("SolidWorks未连接", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                
                ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
                if (swModel == null)
                {
                    MessageBox.Show("没有打开的文档", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                if (swModel.GetType() != (int)SolidWorks.Interop.swconst.swDocumentTypes_e.swDocASSEMBLY)
                {
                    MessageBox.Show("当前文档不是装配体", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                Debug.WriteLine("开始手动刷新BOM数据...");
                infoLabel.Text = "正在刷新数据...";
                refreshButton.Enabled = false;
                
                // 异步调用asm2bom生成BOM
                System.Threading.Tasks.Task.Run(async () =>
                {
                    try
                    {
                        await tools.asm2bom.run(swApp, swModel, "零件", false);
                        Debug.WriteLine("BOM数据刷新完成");
                        
                        // 刷新完成后重新启用按钮
                        this.Invoke(new Action(() =>
                        {
                            refreshButton.Enabled = true;
                            infoLabel.Text = $"零件处理状态监控 (共 {partStatusList.Count} 条记录)";
                        }));
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"刷新BOM失败: {ex.Message}");
                        this.Invoke(new Action(() =>
                        {
                            infoLabel.Text = $"刷新失败: {ex.Message}";
                            refreshButton.Enabled = true;
                        }));
                    }
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"刷新按钮点击失败: {ex.Message}");
                MessageBox.Show($"刷新失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                refreshButton.Enabled = true;
            }
        }
        
        /// <summary>
        /// 单元格点击事件 - 选中零件
        /// </summary>
        private void StatusGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex != 0) return; // 只响应零件名称列的点击
            
            try
            {
                string partName = statusGrid.Rows[e.RowIndex].Cells[0].Value?.ToString();
                if (string.IsNullOrEmpty(partName) || swApp == null) return;
                
                ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
                if (swModel == null || swModel.GetType() != (int)SolidWorks.Interop.swconst.swDocumentTypes_e.swDocASSEMBLY) return;
                
                AssemblyDoc swAssembly = (AssemblyDoc)swModel;
                
                // 遍历组件查找匹配的零件
                object[] components = (object[])swAssembly.GetComponents(false);
                foreach (object compObj in components)
                {
                    Component2 component = (Component2)compObj;
                    string componentName = component.Name2;
                    
                    // 去掉"/"号及之前的文字
                    int slashIndex = componentName.LastIndexOf('/');
                    if (slashIndex >= 0 && slashIndex < componentName.Length - 1)
                    {
                        componentName = componentName.Substring(slashIndex + 1);
                    }
                    
                    // 去掉末尾的"-数字"部分
                    int lastDashIndex = componentName.LastIndexOf('-');
                    if (lastDashIndex > 0 && lastDashIndex < componentName.Length - 1)
                    {
                        string suffix = componentName.Substring(lastDashIndex + 1);
                        if (int.TryParse(suffix, out _))
                        {
                            componentName = componentName.Substring(0, lastDashIndex);
                        }
                    }
                    
                    // 找到匹配的组件并选中
                    if (componentName.Equals(partName, StringComparison.OrdinalIgnoreCase))
                    {
                        component.Select(false);
                        swModel.ViewZoomtofit2();
                        Debug.WriteLine($"已选中零件: {partName}");
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"选中零件失败: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 从 BOM 数据加载零件状态
        /// </summary>
        public void LoadFromBomData(List<PartStatusInfo> bomData)
        {
            partStatusList.Clear();
                    
            foreach (var bomItem in bomData)
            {
                partStatusList.Add(bomItem);
            }
                    
            RefreshStatusDisplay();
        }
        
        /// <summary>
        /// 获取零件数量
        /// </summary>
        public int GetPartCount()
        {
            return partStatusList.Count;
        }
        
        /// <summary>
        /// 更新零件的出图状态
        /// </summary>
        public void UpdatePartDrawnStatus(string partName, string isDrawn)
        {
            var part = partStatusList.Find(p => p.PartName == partName);
            if (part != null)
            {
                part.IsDrawn = isDrawn;
                RefreshStatusDisplay();
                Debug.WriteLine($"已更新零件 '{partName}' 的出图状态为: {isDrawn}");
            }
        }
        
        /// <summary>
        /// 删除零件
        /// </summary>
        public void RemovePart(string partName)
        {
            var part = partStatusList.Find(p => p.PartName == partName);
            if (part != null)
            {
                partStatusList.Remove(part);
                RefreshStatusDisplay();
                Debug.WriteLine($"已从任务窗格移除零件: {partName}");
            }
        }
        
        /// <summary>
        /// 获取指定类型的零件列表
        /// </summary>
        /// <param name="partType">零件类型（如：钣金件、管件、其他）</param>
        /// <returns>符合条件的零件列表</returns>
        public List<PartStatusInfo> GetPartsByType(string partType)
        {
            return partStatusList.FindAll(p => p.PartType == partType);
        }
        
        /// <summary>
        /// 刷新状态显示
        /// </summary>
        private void RefreshStatusDisplay()
        {
            statusGrid.Rows.Clear();
            
            foreach (var status in partStatusList)
            {
                int rowIndex = statusGrid.Rows.Add();
                statusGrid.Rows[rowIndex].Cells[0].Value = status.PartName;
                statusGrid.Rows[rowIndex].Cells[1].Value = status.PartType;
                statusGrid.Rows[rowIndex].Cells[2].Value = status.Dimension;
                statusGrid.Rows[rowIndex].Cells[3].Value = status.IsDrawn;
                statusGrid.Rows[rowIndex].Cells[4].Value = status.Quantity;
                
                // 根据是否出图设置颜色
                if (status.IsDrawn == "已出图")
                {
                    statusGrid.Rows[rowIndex].DefaultCellStyle.BackColor = System.Drawing.Color.LightGreen;
                }
                else
                {
                    statusGrid.Rows[rowIndex].DefaultCellStyle.BackColor = System.Drawing.Color.LightYellow;
                }
            }
            
            infoLabel.Text = $"零件处理状态监控 (共 {partStatusList.Count} 条记录)";
        }

        private DataGridViewTextBoxColumn CreateTextColumn(string name, int width)
        {
            return new DataGridViewTextBoxColumn
            {
                Name = name,
                HeaderText = name,
                Width = width,
                ReadOnly = true
            };
        }

        private string NormalizeComponentName(string componentName)
        {
            if (string.IsNullOrWhiteSpace(componentName))
            {
                return string.Empty;
            }

            // 去掉"/"号及之前的文字
            int slashIndex = componentName.LastIndexOf('/');
            if (slashIndex >= 0 && slashIndex < componentName.Length - 1)
            {
                componentName = componentName.Substring(slashIndex + 1);
            }

            // 去掉末尾的"-数字"实例后缀
            int lastDashIndex = componentName.LastIndexOf('-');
            if (lastDashIndex > 0 && lastDashIndex < componentName.Length - 1)
            {
                string suffix = componentName.Substring(lastDashIndex + 1);
                if (int.TryParse(suffix, out _))
                {
                    componentName = componentName.Substring(0, lastDashIndex);
                }
            }

            return componentName;
        }
    }

}
