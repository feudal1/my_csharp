using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace SolidWorksAddinStudy
{
    [ComVisible(true)]
    [Guid("F95417AA-3B2E-42E0-8B0B-2A2A50E9FA0A")]
    [ProgId("SolidWorksAddinStudy.WorkProjectTaskPaneControl")]
    public class WorkProjectTaskPaneControl : UserControl
    {
        private readonly ListBox projectListBox;
        private readonly TextBox projectNameTextBox;
        private readonly Button addProjectButton;
        private readonly Button removeProjectButton;

        private readonly TextBox rollerTextBox;
        private readonly TextBox pipeTextBox;
        private readonly TextBox sheetMetalTextBox;
        private readonly TextBox machiningTextBox;
        private readonly TextBox purchaseTextBox;
        private readonly TextBox bearingTextBox;
        private readonly TextBox timingBeltTextBox;
        private readonly TextBox folderPathTextBox;
        private readonly Button chooseFolderButton;
        private readonly Button openFolderButton;

        private readonly Label statusLabel;

        private readonly List<WorkProjectItem> projects = new List<WorkProjectItem>();
        private bool isLoadingProject;

        private readonly string saveFilePath;

        public WorkProjectTaskPaneControl()
        {
            saveFilePath = BuildSavePath();

            Dock = DockStyle.Fill;
            MinimumSize = new System.Drawing.Size(420, 480);

            var header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 72,
                Padding = new Padding(8)
            };

            var projectNameLabel = new Label
            {
                Text = "项目名:",
                AutoSize = true,
                Left = 4,
                Top = 12
            };

            projectNameTextBox = new TextBox
            {
                Left = 62,
                Top = 8,
                Width = 170
            };
            projectNameTextBox.KeyDown += ProjectNameTextBox_KeyDown;

            addProjectButton = new Button
            {
                Text = "新建项目",
                Left = 240,
                Top = 6,
                Width = 82,
                Height = 28
            };
            addProjectButton.Click += AddProjectButton_Click;

            removeProjectButton = new Button
            {
                Text = "删除项目",
                Left = 330,
                Top = 6,
                Width = 82,
                Height = 28
            };
            removeProjectButton.Click += RemoveProjectButton_Click;

            statusLabel = new Label
            {
                Text = "工作项目列表",
                AutoSize = true,
                Left = 4,
                Top = 42
            };

            header.Controls.Add(projectNameLabel);
            header.Controls.Add(projectNameTextBox);
            header.Controls.Add(addProjectButton);
            header.Controls.Add(removeProjectButton);
            header.Controls.Add(statusLabel);

            var body = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 150,
                FixedPanel = FixedPanel.Panel1
            };

            projectListBox = new ListBox
            {
                Dock = DockStyle.Fill
            };
            projectListBox.SelectedIndexChanged += ProjectListBox_SelectedIndexChanged;
            body.Panel1.Controls.Add(projectListBox);

            var detailsPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 7,
                Padding = new Padding(8)
            };
            detailsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 86));
            detailsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            var folderPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 70,
                Padding = new Padding(8, 8, 8, 0)
            };

            var folderLabel = new Label
            {
                Text = "项目文件夹:",
                AutoSize = true,
                Left = 0,
                Top = 8
            };
            folderPanel.Controls.Add(folderLabel);

            folderPathTextBox = new TextBox
            {
                Left = 0,
                Top = 28,
                Width = 185
            };
            folderPathTextBox.TextChanged += FolderPathTextBox_TextChanged;
            folderPanel.Controls.Add(folderPathTextBox);

            chooseFolderButton = new Button
            {
                Text = "选择",
                Left = 192,
                Top = 26,
                Width = 52,
                Height = 26
            };
            chooseFolderButton.Click += ChooseFolderButton_Click;
            folderPanel.Controls.Add(chooseFolderButton);

            openFolderButton = new Button
            {
                Text = "打开",
                Left = 248,
                Top = 26,
                Width = 52,
                Height = 26
            };
            openFolderButton.Click += OpenFolderButton_Click;
            folderPanel.Controls.Add(openFolderButton);

            rollerTextBox = CreateDetailTextBox();
            pipeTextBox = CreateDetailTextBox();
            sheetMetalTextBox = CreateDetailTextBox();
            machiningTextBox = CreateDetailTextBox();
            purchaseTextBox = CreateDetailTextBox();
            bearingTextBox = CreateDetailTextBox();
            timingBeltTextBox = CreateDetailTextBox();

            AddDetailRow(detailsPanel, 0, "滚筒出图", rollerTextBox);
            AddDetailRow(detailsPanel, 1, "管件出图", pipeTextBox);
            AddDetailRow(detailsPanel, 2, "钣金出图", sheetMetalTextBox);
            AddDetailRow(detailsPanel, 3, "机加出图", machiningTextBox);
            AddDetailRow(detailsPanel, 4, "外购采购", purchaseTextBox);
            AddDetailRow(detailsPanel, 5, "轴承采购", bearingTextBox);
            AddDetailRow(detailsPanel, 6, "同步带采购", timingBeltTextBox);

            body.Panel2.Controls.Add(detailsPanel);
            body.Panel2.Controls.Add(folderPanel);

            Controls.Add(body);
            Controls.Add(header);

            HookDetailTextChangedEvents();
            LoadProjectsFromLocal();
            RefreshProjectList();
            UpdateDetailAreaEnabledState();
        }

        private static string BuildSavePath()
        {
            string dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "SolidWorksAddinStudy");

            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            return Path.Combine(dir, "work_projects.xml");
        }

        private void AddDetailRow(TableLayoutPanel panel, int rowIndex, string title, TextBox textBox)
        {
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F / panel.RowCount));

            var label = new Label
            {
                Text = title,
                AutoSize = true,
                Anchor = AnchorStyles.Left | AnchorStyles.Top,
                Margin = new Padding(0, 8, 4, 4)
            };

            panel.Controls.Add(label, 0, rowIndex);
            panel.Controls.Add(textBox, 1, rowIndex);
        }

        private static TextBox CreateDetailTextBox()
        {
            return new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical
            };
        }

        private void HookDetailTextChangedEvents()
        {
            rollerTextBox.TextChanged += DetailTextBox_TextChanged;
            pipeTextBox.TextChanged += DetailTextBox_TextChanged;
            sheetMetalTextBox.TextChanged += DetailTextBox_TextChanged;
            machiningTextBox.TextChanged += DetailTextBox_TextChanged;
            purchaseTextBox.TextChanged += DetailTextBox_TextChanged;
            bearingTextBox.TextChanged += DetailTextBox_TextChanged;
            timingBeltTextBox.TextChanged += DetailTextBox_TextChanged;
        }

        private void ProjectNameTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                AddProject();
            }
        }

        private void AddProjectButton_Click(object sender, EventArgs e)
        {
            AddProject();
        }

        private void RemoveProjectButton_Click(object sender, EventArgs e)
        {
            RemoveCurrentProject();
        }

        private void ProjectListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadSelectedProjectToEditor();
        }

        private void FolderPathTextBox_TextChanged(object sender, EventArgs e)
        {
            if (isLoadingProject)
            {
                return;
            }

            var current = GetCurrentProject();
            if (current == null)
            {
                return;
            }

            current.FolderPath = folderPathTextBox.Text?.Trim() ?? string.Empty;
            SaveProjectsToLocal();
            statusLabel.Text = $"已保存: {current.ProjectName}";
        }

        private void DetailTextBox_TextChanged(object sender, EventArgs e)
        {
            if (isLoadingProject)
            {
                return;
            }

            var current = GetCurrentProject();
            if (current == null)
            {
                return;
            }

            current.RollerDrawing = rollerTextBox.Text;
            current.PipeDrawing = pipeTextBox.Text;
            current.SheetMetalDrawing = sheetMetalTextBox.Text;
            current.MachiningDrawing = machiningTextBox.Text;
            current.PurchasedProcurement = purchaseTextBox.Text;
            current.BearingProcurement = bearingTextBox.Text;
            current.TimingBeltProcurement = timingBeltTextBox.Text;

            SaveProjectsToLocal();
            statusLabel.Text = $"已保存: {current.ProjectName}";
        }

        private void ChooseFolderButton_Click(object sender, EventArgs e)
        {
            var current = GetCurrentProject();
            if (current == null)
            {
                return;
            }

            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "选择项目文件夹";
                if (!string.IsNullOrWhiteSpace(current.FolderPath) && Directory.Exists(current.FolderPath))
                {
                    dialog.SelectedPath = current.FolderPath;
                }

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    folderPathTextBox.Text = dialog.SelectedPath;
                }
            }
        }

        private void OpenFolderButton_Click(object sender, EventArgs e)
        {
            var current = GetCurrentProject();
            if (current == null)
            {
                return;
            }

            string path = (current.FolderPath ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(path))
            {
                MessageBox.Show("请先设置项目文件夹地址", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (!Directory.Exists(path))
            {
                MessageBox.Show("项目文件夹不存在，请检查路径", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                Process.Start("explorer.exe", $"\"{path}\"");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开文件夹失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AddProject()
        {
            string projectName = (projectNameTextBox.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(projectName))
            {
                MessageBox.Show("请输入项目名", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (projects.Exists(p => string.Equals(p.ProjectName, projectName, StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show("该项目已存在，请使用不同名称", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var item = new WorkProjectItem
            {
                ProjectName = projectName
            };
            projects.Add(item);

            SaveProjectsToLocal();
            RefreshProjectList();

            projectListBox.SelectedItem = projectName;
            projectNameTextBox.Clear();
            statusLabel.Text = $"已新建项目: {projectName}";
        }

        private void RemoveCurrentProject()
        {
            var current = GetCurrentProject();
            if (current == null)
            {
                return;
            }

            var result = MessageBox.Show(
                $"确认删除项目“{current.ProjectName}”？",
                "删除确认",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result != DialogResult.Yes)
            {
                return;
            }

            projects.Remove(current);
            SaveProjectsToLocal();
            RefreshProjectList();

            statusLabel.Text = "项目已删除";
        }

        private WorkProjectItem GetCurrentProject()
        {
            if (projectListBox.SelectedIndex < 0 || projectListBox.SelectedIndex >= projects.Count)
            {
                return null;
            }

            return projects[projectListBox.SelectedIndex];
        }

        private void LoadSelectedProjectToEditor()
        {
            isLoadingProject = true;
            try
            {
                var current = GetCurrentProject();
                if (current == null)
                {
                    folderPathTextBox.Text = string.Empty;
                    rollerTextBox.Text = string.Empty;
                    pipeTextBox.Text = string.Empty;
                    sheetMetalTextBox.Text = string.Empty;
                    machiningTextBox.Text = string.Empty;
                    purchaseTextBox.Text = string.Empty;
                    bearingTextBox.Text = string.Empty;
                    timingBeltTextBox.Text = string.Empty;
                    statusLabel.Text = "请选择或新建项目";
                    return;
                }

                folderPathTextBox.Text = current.FolderPath ?? string.Empty;
                rollerTextBox.Text = current.RollerDrawing ?? string.Empty;
                pipeTextBox.Text = current.PipeDrawing ?? string.Empty;
                sheetMetalTextBox.Text = current.SheetMetalDrawing ?? string.Empty;
                machiningTextBox.Text = current.MachiningDrawing ?? string.Empty;
                purchaseTextBox.Text = current.PurchasedProcurement ?? string.Empty;
                bearingTextBox.Text = current.BearingProcurement ?? string.Empty;
                timingBeltTextBox.Text = current.TimingBeltProcurement ?? string.Empty;
                statusLabel.Text = $"当前项目: {current.ProjectName}";
            }
            finally
            {
                isLoadingProject = false;
                UpdateDetailAreaEnabledState();
            }
        }

        private void UpdateDetailAreaEnabledState()
        {
            bool enabled = projectListBox.SelectedIndex >= 0;
            folderPathTextBox.Enabled = enabled;
            chooseFolderButton.Enabled = enabled;
            openFolderButton.Enabled = enabled;
            rollerTextBox.Enabled = enabled;
            pipeTextBox.Enabled = enabled;
            sheetMetalTextBox.Enabled = enabled;
            machiningTextBox.Enabled = enabled;
            purchaseTextBox.Enabled = enabled;
            bearingTextBox.Enabled = enabled;
            timingBeltTextBox.Enabled = enabled;
            removeProjectButton.Enabled = enabled;
        }

        private void RefreshProjectList()
        {
            int selectedIndex = projectListBox.SelectedIndex;

            projectListBox.BeginUpdate();
            try
            {
                projectListBox.Items.Clear();
                foreach (var project in projects)
                {
                    projectListBox.Items.Add(project.ProjectName);
                }
            }
            finally
            {
                projectListBox.EndUpdate();
            }

            if (projectListBox.Items.Count == 0)
            {
                projectListBox.SelectedIndex = -1;
                UpdateDetailAreaEnabledState();
                statusLabel.Text = "暂无项目，请先新建";
                return;
            }

            if (selectedIndex < 0 || selectedIndex >= projectListBox.Items.Count)
            {
                selectedIndex = 0;
            }

            projectListBox.SelectedIndex = selectedIndex;
        }

        private void SaveProjectsToLocal()
        {
            try
            {
                var serializer = new XmlSerializer(typeof(WorkProjectStore));
                var store = new WorkProjectStore { Projects = projects };

                using (var fs = new FileStream(saveFilePath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    serializer.Serialize(fs, store);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"保存工作项目失败: {ex.Message}");
            }
        }

        private void LoadProjectsFromLocal()
        {
            try
            {
                if (!File.Exists(saveFilePath))
                {
                    return;
                }

                var serializer = new XmlSerializer(typeof(WorkProjectStore));
                using (var fs = new FileStream(saveFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    var store = serializer.Deserialize(fs) as WorkProjectStore;
                    if (store?.Projects != null)
                    {
                        projects.Clear();
                        projects.AddRange(store.Projects);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"读取工作项目失败: {ex.Message}");
            }
        }
    }

    [Serializable]
    public class WorkProjectStore
    {
        public List<WorkProjectItem> Projects { get; set; } = new List<WorkProjectItem>();
    }

    [Serializable]
    public class WorkProjectItem
    {
        public string ProjectName { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
        public string RollerDrawing { get; set; } = string.Empty;
        public string PipeDrawing { get; set; } = string.Empty;
        public string SheetMetalDrawing { get; set; } = string.Empty;
        public string MachiningDrawing { get; set; } = string.Empty;
        public string PurchasedProcurement { get; set; } = string.Empty;
        public string BearingProcurement { get; set; } = string.Empty;
        public string TimingBeltProcurement { get; set; } = string.Empty;
    }
}
