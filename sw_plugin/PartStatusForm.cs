using System;
using System.Windows.Forms;
using System.Collections.Generic;
using SolidWorks.Interop.sldworks;
using System.Diagnostics;
using System.IO;

namespace SolidWorksAddinStudy
{
    /// <summary>
    /// 零件处理状态任务窗格
    /// </summary>
    public partial class PartStatusForm : Form
    {
        private DataGridView statusGrid;
        private Label infoLabel;
        
        // 存储零件状态数据
        private List<PartStatusInfo> partStatusList = new List<PartStatusInfo>();
        
        // SolidWorks应用实例
        private SldWorks swApp;
        
        public PartStatusForm(SldWorks swApp)
        {
            this.swApp = swApp;
            InitializeComponent();
        }
        
        private void InitializeComponent()
        {
            this.Text = "零件处理状态";
            this.Size = new System.Drawing.Size(600, 600);
            this.MinimumSize = new System.Drawing.Size(400, 400);
            
            // 创建信息标签
            infoLabel = new Label();
            infoLabel.Text = "零件处理状态监控";
            infoLabel.Location = new System.Drawing.Point(10, 10);
            infoLabel.Size = new System.Drawing.Size(300, 20);
            infoLabel.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Bold);
            
            // 创建DataGridView显示状态
            statusGrid = new DataGridView();
            statusGrid.Location = new System.Drawing.Point(10, 40);
            statusGrid.Size = new System.Drawing.Size(570, 510);
            statusGrid.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            statusGrid.AllowUserToAddRows = false;
            statusGrid.AllowUserToDeleteRows = false;
            statusGrid.ReadOnly = true;
            statusGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            statusGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            statusGrid.MultiSelect = false;
            
            // 设置列
            statusGrid.ColumnCount = 5;
            statusGrid.Columns[0].Name = "零件名称";
            statusGrid.Columns[1].Name = "零件类型";
            statusGrid.Columns[2].Name = "规格尺寸";
            statusGrid.Columns[3].Name = "是否出图";
            statusGrid.Columns[4].Name = "数量";
            
            // 设置列宽比例
            statusGrid.Columns[0].Width = 150;
            statusGrid.Columns[1].Width = 80;
            statusGrid.Columns[2].Width = 120;
            statusGrid.Columns[3].Width = 80;
            statusGrid.Columns[4].Width = 60;
            
            // 添加单元格点击事件
            statusGrid.CellClick += StatusGrid_CellClick;
            
            // 添加控件
            this.Controls.Add(infoLabel);
            this.Controls.Add(statusGrid);
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
        
        /// <summary>
        /// 获取当前SolidWorks文档的零件名称
        /// </summary>
        public static string GetCurrentPartName(SldWorks swApp)
        {
            try
            {
                ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
                if (swModel != null)
                {
                    string fullPath = swModel.GetPathName();
                    if (!string.IsNullOrEmpty(fullPath))
                    {
                        return System.IO.Path.GetFileNameWithoutExtension(fullPath);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"获取零件名称失败: {ex.Message}");
            }
            return "未知零件";
        }
    }
    
    /// <summary>
    /// 零件状态信息类
    /// </summary>
    public class PartStatusInfo
    {
        public string PartName { get; set; }
        public string PartType { get; set; }
        public string Dimension { get; set; }
        public string IsDrawn { get; set; }
        public string Quantity { get; set; }
    }
}
