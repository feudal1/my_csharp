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

        private readonly TextBox rollerOutputTextBox;
        private readonly TextBox rollerFollowUpTextBox;
        private readonly TextBox pipeOutputTextBox;
        private readonly TextBox pipeFollowUpTextBox;
        private readonly TextBox sheetMetalOutputTextBox;
        private readonly TextBox sheetMetalFollowUpTextBox;
        private readonly TextBox machiningOutputTextBox;
        private readonly TextBox machiningFollowUpTextBox;
        private readonly TextBox purchaseOutputTextBox;
        private readonly TextBox purchaseFollowUpTextBox;
        private readonly TextBox bearingOutputTextBox;
        private readonly TextBox bearingFollowUpTextBox;
        private readonly TextBox timingBeltOutputTextBox;
        private readonly TextBox timingBeltFollowUpTextBox;
        private readonly TextBox countersinkMarkingOutputTextBox;
        private readonly TextBox countersinkMarkingFollowUpTextBox;
        private readonly TextBox folderPathTextBox;
        private readonly Button chooseFolderButton;
        private readonly Button openFolderButton;

        private readonly Label statusLabel;

        private readonly List<WorkProjectItem> projects = new List<WorkProjectItem>();
        private bool isLoadingProject;

        private readonly string saveFilePath;
        private static readonly TimeSpan FollowUpReminderDelay = TimeSpan.FromDays(2);

        public WorkProjectTaskPaneControl()
        {
            saveFilePath = BuildSavePath();

            Dock = DockStyle.Fill;
            AutoScaleMode = AutoScaleMode.Dpi;
            MinimumSize = new System.Drawing.Size(320, 240);

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
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 3,
                RowCount = 9,
                Padding = new Padding(8)
            };
            detailsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 86));
            detailsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            detailsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));

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

            detailsPanel.Controls.Add(new Label
            {
                Text = "",
                AutoSize = true,
                Margin = new Padding(0, 0, 4, 6)
            }, 0, 0);
            detailsPanel.Controls.Add(new Label
            {
                Text = "出图",
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 0, 4, 6)
            }, 1, 0);
            detailsPanel.Controls.Add(new Label
            {
                Text = "跟进",
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 0, 4, 6)
            }, 2, 0);

            rollerOutputTextBox = CreateDetailTextBox();
            rollerFollowUpTextBox = CreateDetailTextBox();
            pipeOutputTextBox = CreateDetailTextBox();
            pipeFollowUpTextBox = CreateDetailTextBox();
            sheetMetalOutputTextBox = CreateDetailTextBox();
            sheetMetalFollowUpTextBox = CreateDetailTextBox();
            machiningOutputTextBox = CreateDetailTextBox();
            machiningFollowUpTextBox = CreateDetailTextBox();
            purchaseOutputTextBox = CreateDetailTextBox();
            purchaseFollowUpTextBox = CreateDetailTextBox();
            bearingOutputTextBox = CreateDetailTextBox();
            bearingFollowUpTextBox = CreateDetailTextBox();
            timingBeltOutputTextBox = CreateDetailTextBox();
            timingBeltFollowUpTextBox = CreateDetailTextBox();
            countersinkMarkingOutputTextBox = CreateDetailTextBox();
            countersinkMarkingFollowUpTextBox = CreateDetailTextBox();

            AddDetailRow(detailsPanel, 1, "滚筒出图", rollerOutputTextBox, rollerFollowUpTextBox);
            AddDetailRow(detailsPanel, 2, "管件出图", pipeOutputTextBox, pipeFollowUpTextBox);
            AddDetailRow(detailsPanel, 3, "钣金出图", sheetMetalOutputTextBox, sheetMetalFollowUpTextBox);
            AddDetailRow(detailsPanel, 4, "机加出图", machiningOutputTextBox, machiningFollowUpTextBox);
            AddDetailRow(detailsPanel, 5, "外购采购", purchaseOutputTextBox, purchaseFollowUpTextBox);
            AddDetailRow(detailsPanel, 6, "轴承采购", bearingOutputTextBox, bearingFollowUpTextBox);
            AddDetailRow(detailsPanel, 7, "同步带采购", timingBeltOutputTextBox, timingBeltFollowUpTextBox);
            AddDetailRow(detailsPanel, 8, "打标沉孔", countersinkMarkingOutputTextBox, countersinkMarkingFollowUpTextBox);

            var detailsScrollPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true
            };
            detailsScrollPanel.Controls.Add(detailsPanel);
            detailsScrollPanel.Resize += (sender, args) => RefreshDetailsScrollRange(detailsScrollPanel, detailsPanel);
            detailsPanel.SizeChanged += (sender, args) => RefreshDetailsScrollRange(detailsScrollPanel, detailsPanel);

            body.Panel2.Controls.Add(detailsScrollPanel);
            body.Panel2.Controls.Add(folderPanel);

            Controls.Add(body);
            Controls.Add(header);

            HookDetailTextChangedEvents();
            LoadProjectsFromLocal();
            RefreshProjectList();
            UpdateDetailAreaEnabledState();
            RefreshDetailsScrollRange(detailsScrollPanel, detailsPanel);
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

        private void AddDetailRow(TableLayoutPanel panel, int rowIndex, string title, TextBox outputTextBox, TextBox followUpTextBox)
        {
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 84F));

            var label = new Label
            {
                Text = title,
                AutoSize = true,
                Anchor = AnchorStyles.Left | AnchorStyles.Top,
                Margin = new Padding(0, 8, 4, 4)
            };

            panel.Controls.Add(label, 0, rowIndex);
            panel.Controls.Add(outputTextBox, 1, rowIndex);
            panel.Controls.Add(followUpTextBox, 2, rowIndex);
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

        private static void RefreshDetailsScrollRange(Panel scrollPanel, TableLayoutPanel detailsPanel)
        {
            int contentHeight = detailsPanel.PreferredSize.Height + 8;
            scrollPanel.AutoScrollMinSize = new System.Drawing.Size(0, contentHeight);
        }

        private void HookDetailTextChangedEvents()
        {
            rollerOutputTextBox.TextChanged += DetailTextBox_TextChanged;
            rollerFollowUpTextBox.TextChanged += DetailTextBox_TextChanged;
            pipeOutputTextBox.TextChanged += DetailTextBox_TextChanged;
            pipeFollowUpTextBox.TextChanged += DetailTextBox_TextChanged;
            sheetMetalOutputTextBox.TextChanged += DetailTextBox_TextChanged;
            sheetMetalFollowUpTextBox.TextChanged += DetailTextBox_TextChanged;
            machiningOutputTextBox.TextChanged += DetailTextBox_TextChanged;
            machiningFollowUpTextBox.TextChanged += DetailTextBox_TextChanged;
            purchaseOutputTextBox.TextChanged += DetailTextBox_TextChanged;
            purchaseFollowUpTextBox.TextChanged += DetailTextBox_TextChanged;
            bearingOutputTextBox.TextChanged += DetailTextBox_TextChanged;
            bearingFollowUpTextBox.TextChanged += DetailTextBox_TextChanged;
            timingBeltOutputTextBox.TextChanged += DetailTextBox_TextChanged;
            timingBeltFollowUpTextBox.TextChanged += DetailTextBox_TextChanged;
            countersinkMarkingOutputTextBox.TextChanged += DetailTextBox_TextChanged;
            countersinkMarkingFollowUpTextBox.TextChanged += DetailTextBox_TextChanged;
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

            UpdateDrawingAndFollowUp(current.RollerDrawing, rollerOutputTextBox.Text, rollerFollowUpTextBox.Text, t => current.RollerDrawingUpdatedAt = t, out string rollerDrawing, out string rollerFollowUp);
            current.RollerDrawing = rollerDrawing;
            current.RollerFollowUp = rollerFollowUp;

            UpdateDrawingAndFollowUp(current.PipeDrawing, pipeOutputTextBox.Text, pipeFollowUpTextBox.Text, t => current.PipeDrawingUpdatedAt = t, out string pipeDrawing, out string pipeFollowUp);
            current.PipeDrawing = pipeDrawing;
            current.PipeFollowUp = pipeFollowUp;

            UpdateDrawingAndFollowUp(current.SheetMetalDrawing, sheetMetalOutputTextBox.Text, sheetMetalFollowUpTextBox.Text, t => current.SheetMetalDrawingUpdatedAt = t, out string sheetMetalDrawing, out string sheetMetalFollowUp);
            current.SheetMetalDrawing = sheetMetalDrawing;
            current.SheetMetalFollowUp = sheetMetalFollowUp;

            UpdateDrawingAndFollowUp(current.MachiningDrawing, machiningOutputTextBox.Text, machiningFollowUpTextBox.Text, t => current.MachiningDrawingUpdatedAt = t, out string machiningDrawing, out string machiningFollowUp);
            current.MachiningDrawing = machiningDrawing;
            current.MachiningFollowUp = machiningFollowUp;

            UpdateDrawingAndFollowUp(current.PurchasedProcurement, purchaseOutputTextBox.Text, purchaseFollowUpTextBox.Text, t => current.PurchasedProcurementUpdatedAt = t, out string purchasedProcurement, out string purchasedFollowUp);
            current.PurchasedProcurement = purchasedProcurement;
            current.PurchasedFollowUp = purchasedFollowUp;

            UpdateDrawingAndFollowUp(current.BearingProcurement, bearingOutputTextBox.Text, bearingFollowUpTextBox.Text, t => current.BearingProcurementUpdatedAt = t, out string bearingProcurement, out string bearingFollowUp);
            current.BearingProcurement = bearingProcurement;
            current.BearingFollowUp = bearingFollowUp;

            UpdateDrawingAndFollowUp(current.TimingBeltProcurement, timingBeltOutputTextBox.Text, timingBeltFollowUpTextBox.Text, t => current.TimingBeltProcurementUpdatedAt = t, out string timingBeltProcurement, out string timingBeltFollowUp);
            current.TimingBeltProcurement = timingBeltProcurement;
            current.TimingBeltFollowUp = timingBeltFollowUp;

            UpdateDrawingAndFollowUp(current.CountersinkMarkingDrawing, countersinkMarkingOutputTextBox.Text, countersinkMarkingFollowUpTextBox.Text, t => current.CountersinkMarkingDrawingUpdatedAt = t, out string countersinkMarkingDrawing, out string countersinkMarkingFollowUp);
            current.CountersinkMarkingDrawing = countersinkMarkingDrawing;
            current.CountersinkMarkingFollowUp = countersinkMarkingFollowUp;

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
                    rollerOutputTextBox.Text = string.Empty;
                    rollerFollowUpTextBox.Text = string.Empty;
                    pipeOutputTextBox.Text = string.Empty;
                    pipeFollowUpTextBox.Text = string.Empty;
                    sheetMetalOutputTextBox.Text = string.Empty;
                    sheetMetalFollowUpTextBox.Text = string.Empty;
                    machiningOutputTextBox.Text = string.Empty;
                    machiningFollowUpTextBox.Text = string.Empty;
                    purchaseOutputTextBox.Text = string.Empty;
                    purchaseFollowUpTextBox.Text = string.Empty;
                    bearingOutputTextBox.Text = string.Empty;
                    bearingFollowUpTextBox.Text = string.Empty;
                    timingBeltOutputTextBox.Text = string.Empty;
                    timingBeltFollowUpTextBox.Text = string.Empty;
                    countersinkMarkingOutputTextBox.Text = string.Empty;
                    countersinkMarkingFollowUpTextBox.Text = string.Empty;
                    statusLabel.Text = "请选择或新建项目";
                    return;
                }

                folderPathTextBox.Text = current.FolderPath ?? string.Empty;
                rollerOutputTextBox.Text = current.RollerDrawing ?? string.Empty;
                rollerFollowUpTextBox.Text = current.RollerFollowUp ?? string.Empty;
                pipeOutputTextBox.Text = current.PipeDrawing ?? string.Empty;
                pipeFollowUpTextBox.Text = current.PipeFollowUp ?? string.Empty;
                sheetMetalOutputTextBox.Text = current.SheetMetalDrawing ?? string.Empty;
                sheetMetalFollowUpTextBox.Text = current.SheetMetalFollowUp ?? string.Empty;
                machiningOutputTextBox.Text = current.MachiningDrawing ?? string.Empty;
                machiningFollowUpTextBox.Text = current.MachiningFollowUp ?? string.Empty;
                purchaseOutputTextBox.Text = current.PurchasedProcurement ?? string.Empty;
                purchaseFollowUpTextBox.Text = current.PurchasedFollowUp ?? string.Empty;
                bearingOutputTextBox.Text = current.BearingProcurement ?? string.Empty;
                bearingFollowUpTextBox.Text = current.BearingFollowUp ?? string.Empty;
                timingBeltOutputTextBox.Text = current.TimingBeltProcurement ?? string.Empty;
                timingBeltFollowUpTextBox.Text = current.TimingBeltFollowUp ?? string.Empty;
                countersinkMarkingOutputTextBox.Text = current.CountersinkMarkingDrawing ?? string.Empty;
                countersinkMarkingFollowUpTextBox.Text = current.CountersinkMarkingFollowUp ?? string.Empty;
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
            rollerOutputTextBox.Enabled = enabled;
            rollerFollowUpTextBox.Enabled = enabled;
            pipeOutputTextBox.Enabled = enabled;
            pipeFollowUpTextBox.Enabled = enabled;
            sheetMetalOutputTextBox.Enabled = enabled;
            sheetMetalFollowUpTextBox.Enabled = enabled;
            machiningOutputTextBox.Enabled = enabled;
            machiningFollowUpTextBox.Enabled = enabled;
            purchaseOutputTextBox.Enabled = enabled;
            purchaseFollowUpTextBox.Enabled = enabled;
            bearingOutputTextBox.Enabled = enabled;
            bearingFollowUpTextBox.Enabled = enabled;
            timingBeltOutputTextBox.Enabled = enabled;
            timingBeltFollowUpTextBox.Enabled = enabled;
            countersinkMarkingOutputTextBox.Enabled = enabled;
            countersinkMarkingFollowUpTextBox.Enabled = enabled;
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
                        NormalizeLegacyReminderTimestamps();
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"读取工作项目失败: {ex.Message}");
            }
        }

        public void PromptFollowUpRemindersAtStartup()
        {
            var dueItems = CollectDueFollowUps();
            if (dueItems.Count == 0)
            {
                return;
            }

            string message = "以下项目出图已超过2天且尚未填写跟进，是否现在跟进？\n\n"
                           + string.Join("\n", dueItems);

            DialogResult result = MessageBox.Show(
                message,
                "工作项目跟进提醒",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                FocusProjectByName(dueItems[0].ProjectName);
            }
        }

        private void UpdateDrawingAndFollowUp(
            string oldDrawingValue,
            string newDrawingValue,
            string newFollowUpValue,
            Action<DateTime> setUpdatedAt,
            out string drawingValue,
            out string followUpValue)
        {
            drawingValue = newDrawingValue ?? string.Empty;
            followUpValue = newFollowUpValue ?? string.Empty;

            if (!string.Equals((oldDrawingValue ?? string.Empty), drawingValue, StringComparison.Ordinal))
            {
                if (string.IsNullOrWhiteSpace(drawingValue))
                {
                    setUpdatedAt(DateTime.MinValue);
                }
                else
                {
                    setUpdatedAt(DateTime.Now);
                }
            }
        }

        private void NormalizeLegacyReminderTimestamps()
        {
            bool changed = false;
            DateTime now = DateTime.Now;

            foreach (var item in projects)
            {
                changed |= EnsureLegacyTimestamp(item.RollerDrawing, item.RollerDrawingUpdatedAt, t => item.RollerDrawingUpdatedAt = t, now);
                changed |= EnsureLegacyTimestamp(item.PipeDrawing, item.PipeDrawingUpdatedAt, t => item.PipeDrawingUpdatedAt = t, now);
                changed |= EnsureLegacyTimestamp(item.SheetMetalDrawing, item.SheetMetalDrawingUpdatedAt, t => item.SheetMetalDrawingUpdatedAt = t, now);
                changed |= EnsureLegacyTimestamp(item.MachiningDrawing, item.MachiningDrawingUpdatedAt, t => item.MachiningDrawingUpdatedAt = t, now);
                changed |= EnsureLegacyTimestamp(item.PurchasedProcurement, item.PurchasedProcurementUpdatedAt, t => item.PurchasedProcurementUpdatedAt = t, now);
                changed |= EnsureLegacyTimestamp(item.BearingProcurement, item.BearingProcurementUpdatedAt, t => item.BearingProcurementUpdatedAt = t, now);
                changed |= EnsureLegacyTimestamp(item.TimingBeltProcurement, item.TimingBeltProcurementUpdatedAt, t => item.TimingBeltProcurementUpdatedAt = t, now);
                changed |= EnsureLegacyTimestamp(item.CountersinkMarkingDrawing, item.CountersinkMarkingDrawingUpdatedAt, t => item.CountersinkMarkingDrawingUpdatedAt = t, now);
            }

            if (changed)
            {
                SaveProjectsToLocal();
            }
        }

        private static bool EnsureLegacyTimestamp(string drawingValue, DateTime timestamp, Action<DateTime> setter, DateTime now)
        {
            if (string.IsNullOrWhiteSpace(drawingValue) || timestamp > DateTime.MinValue)
            {
                return false;
            }

            setter(now);
            return true;
        }

        private List<FollowUpDueItem> CollectDueFollowUps()
        {
            var dueItems = new List<FollowUpDueItem>();
            DateTime now = DateTime.Now;

            foreach (var project in projects)
            {
                AddDueItemIfNeeded(dueItems, project.ProjectName, "滚筒出图", project.RollerDrawing, project.RollerFollowUp, project.RollerDrawingUpdatedAt, now);
                AddDueItemIfNeeded(dueItems, project.ProjectName, "管件出图", project.PipeDrawing, project.PipeFollowUp, project.PipeDrawingUpdatedAt, now);
                AddDueItemIfNeeded(dueItems, project.ProjectName, "钣金出图", project.SheetMetalDrawing, project.SheetMetalFollowUp, project.SheetMetalDrawingUpdatedAt, now);
                AddDueItemIfNeeded(dueItems, project.ProjectName, "机加出图", project.MachiningDrawing, project.MachiningFollowUp, project.MachiningDrawingUpdatedAt, now);
                AddDueItemIfNeeded(dueItems, project.ProjectName, "外购采购", project.PurchasedProcurement, project.PurchasedFollowUp, project.PurchasedProcurementUpdatedAt, now);
                AddDueItemIfNeeded(dueItems, project.ProjectName, "轴承采购", project.BearingProcurement, project.BearingFollowUp, project.BearingProcurementUpdatedAt, now);
                AddDueItemIfNeeded(dueItems, project.ProjectName, "同步带采购", project.TimingBeltProcurement, project.TimingBeltFollowUp, project.TimingBeltProcurementUpdatedAt, now);
                AddDueItemIfNeeded(dueItems, project.ProjectName, "打标沉孔", project.CountersinkMarkingDrawing, project.CountersinkMarkingFollowUp, project.CountersinkMarkingDrawingUpdatedAt, now);
            }

            return dueItems;
        }

        private void AddDueItemIfNeeded(
            List<FollowUpDueItem> dueItems,
            string projectName,
            string fieldTitle,
            string drawingValue,
            string followUpValue,
            DateTime drawingUpdatedAt,
            DateTime now)
        {
            if (string.IsNullOrWhiteSpace(drawingValue))
            {
                return;
            }

            if (!string.IsNullOrWhiteSpace(followUpValue))
            {
                return;
            }

            if (drawingUpdatedAt <= DateTime.MinValue)
            {
                return;
            }

            if ((now - drawingUpdatedAt) < FollowUpReminderDelay)
            {
                return;
            }

            dueItems.Add(new FollowUpDueItem
            {
                ProjectName = projectName,
                Label = fieldTitle
            });
        }

        private void FocusProjectByName(string projectName)
        {
            if (string.IsNullOrWhiteSpace(projectName))
            {
                return;
            }

            for (int i = 0; i < projects.Count; i++)
            {
                if (string.Equals(projects[i].ProjectName, projectName, StringComparison.OrdinalIgnoreCase))
                {
                    projectListBox.SelectedIndex = i;
                    return;
                }
            }
        }
    }

    internal sealed class FollowUpDueItem
    {
        public string ProjectName { get; set; } = string.Empty;
        public string Label { get; set; } = string.Empty;

        public override string ToString()
        {
            return $"{ProjectName} - {Label}";
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
        public DateTime RollerDrawingUpdatedAt { get; set; } = DateTime.MinValue;
        public string RollerFollowUp { get; set; } = string.Empty;
        public string PipeDrawing { get; set; } = string.Empty;
        public DateTime PipeDrawingUpdatedAt { get; set; } = DateTime.MinValue;
        public string PipeFollowUp { get; set; } = string.Empty;
        public string SheetMetalDrawing { get; set; } = string.Empty;
        public DateTime SheetMetalDrawingUpdatedAt { get; set; } = DateTime.MinValue;
        public string SheetMetalFollowUp { get; set; } = string.Empty;
        public string MachiningDrawing { get; set; } = string.Empty;
        public DateTime MachiningDrawingUpdatedAt { get; set; } = DateTime.MinValue;
        public string MachiningFollowUp { get; set; } = string.Empty;
        public string PurchasedProcurement { get; set; } = string.Empty;
        public DateTime PurchasedProcurementUpdatedAt { get; set; } = DateTime.MinValue;
        public string PurchasedFollowUp { get; set; } = string.Empty;
        public string BearingProcurement { get; set; } = string.Empty;
        public DateTime BearingProcurementUpdatedAt { get; set; } = DateTime.MinValue;
        public string BearingFollowUp { get; set; } = string.Empty;
        public string TimingBeltProcurement { get; set; } = string.Empty;
        public DateTime TimingBeltProcurementUpdatedAt { get; set; } = DateTime.MinValue;
        public string TimingBeltFollowUp { get; set; } = string.Empty;
        public string CountersinkMarkingDrawing { get; set; } = string.Empty;
        public DateTime CountersinkMarkingDrawingUpdatedAt { get; set; } = DateTime.MinValue;
        public string CountersinkMarkingFollowUp { get; set; } = string.Empty;
    }
}
