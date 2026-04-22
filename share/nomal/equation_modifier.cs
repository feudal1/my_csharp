using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace tools
{
    /// <summary>
    /// 方程式管理器：支持单零件或多零件统一窗口；列表项带零件前缀以区分重名变量。
    /// </summary>
    public class EquationModifier
    {
        private sealed class EquationListRow
        {
            public ModelDoc2 Part = null!;
            public int EquationIndex;
            public string PartDisplayLabel = "";
        }

        /// <summary>
        /// 单零件入口（兼容旧调用）。
        /// </summary>
        public static string ModifyEquations(SldWorks swApp, ModelDoc2 swModel)
        {
            return ModifyEquationsForParts(swApp, new List<ModelDoc2> { swModel });
        }

        /// <summary>
        /// 多零件：一个窗口列出所有方程式，每项带 [零件标签] 前缀。
        /// </summary>
        public static string ModifyEquationsForParts(SldWorks swApp, IList<ModelDoc2> parts)
        {
            if (swApp == null || parts == null || parts.Count == 0)
            {
                return "SolidWorks 或零件列表无效";
            }

            var unique = DeduplicateParts(parts);
            if (unique.Count == 0)
            {
                return "没有可编辑的零件";
            }

            var visibleParts = new List<ModelDoc2>();
            var invisiblePartNames = new List<string>();

            foreach (var p in unique)
            {
                try
                {
                    p.Visible = true;
                    if (p.Visible)
                    {
                        visibleParts.Add(p);
                    }
                    else
                    {
                        invisiblePartNames.Add(p.GetTitle());
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"设置零件可见失败: {ex.Message}");
                    invisiblePartNames.Add(p.GetTitle());
                }
            }

            try
            {
                if (visibleParts.Count == 0)
                {
                    swApp.SendMsgToUser("所选零件文档均不可见，无法修改方程式");
                    return "所选零件文档均不可见，无法修改方程式";
                }

                if (invisiblePartNames.Count > 0)
                {
                    swApp.SendMsgToUser($"以下零件不可见，已跳过：{string.Join("、", invisiblePartNames.Distinct(StringComparer.OrdinalIgnoreCase))}");
                }

                var rows = BuildEquationRows(visibleParts);
                if (rows.Count == 0)
                {
                    swApp.SendMsgToUser("所选零件中均没有方程式");
                    return "所选零件中均没有方程式";
                }

                var applied = new[] { false };
                if (!ShowUnifiedEquationDialog(swApp, rows, applied))
                {
                    return "用户取消了操作";
                }

                return applied[0] ? "方程式修改完成" : "未应用任何修改";
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"修改方程式失败：{ex.Message}");
                swApp?.SendMsgToUser($"修改方程式失败：{ex.Message}");
                return $"修改方程式失败：{ex.Message}";
            }
            finally
            {
                foreach (var p in visibleParts)
                {
                    try
                    {
                        p.Visible = false;
                    }
                    catch
                    {
                        // ignore
                    }
                }
            }
        }

        private static List<ModelDoc2> DeduplicateParts(IEnumerable<ModelDoc2> parts)
        {
            var list = new List<ModelDoc2>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var p in parts)
            {
                if (p == null || p.GetType() != (int)swDocumentTypes_e.swDocPART)
                {
                    continue;
                }

                string key = p.GetPathName();
                if (string.IsNullOrEmpty(key))
                {
                    key = p.GetTitle();
                }

                if (seen.Add(key))
                {
                    list.Add(p);
                }
            }

            return list;
        }

        private static string GetPartDisplayLabel(ModelDoc2 p, IReadOnlyList<ModelDoc2> allParts)
        {
            string title = p.GetTitle();
            int sameTitle = allParts.Count(x => string.Equals(x.GetTitle(), title, StringComparison.OrdinalIgnoreCase));
            if (sameTitle > 1)
            {
                string path = p.GetPathName();
                return string.IsNullOrEmpty(path)
                    ? title
                    : $"{Path.GetFileName(path)}  |  {path}";
            }

            return title;
        }

        private static List<EquationListRow> BuildEquationRows(List<ModelDoc2> parts)
        {
            var rows = new List<EquationListRow>();

            foreach (var p in parts)
            {
                EquationMgr eqnMgr = null;
                try
                {
                    eqnMgr = (EquationMgr)p.GetEquationMgr();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"获取方程式管理器失败 {p.GetTitle()}: {ex.Message}");
                }

                if (eqnMgr == null)
                {
                    continue;
                }

                int count = eqnMgr.GetCount();
                if (count <= 0)
                {
                    continue;
                }

                string partLabel = GetPartDisplayLabel(p, parts);

                for (int i = 0; i < count; i++)
                {
                    rows.Add(new EquationListRow
                    {
                        Part = p,
                        EquationIndex = i,
                        PartDisplayLabel = partLabel
                    });
                }
            }

            return rows;
        }

        private static string FormatListLine(EquationListRow row, EquationMgr mgr)
        {
            string equation = mgr.get_Equation(row.EquationIndex);
            bool isGlobalVar = mgr.get_GlobalVariable(row.EquationIndex);
            string typeStr = isGlobalVar ? "[全局变量]" : "[尺寸方程]";
            return $"[零件: {row.PartDisplayLabel}]  #{row.EquationIndex + 1} {typeStr} {equation}";
        }

        private static bool ShowUnifiedEquationDialog(SldWorks swApp, List<EquationListRow> rows, bool[] appliedFlag)
        {
            using (var form = new Form())
            {
                int partCount = rows
                    .Select(r =>
                    {
                        string k = r.Part.GetPathName();
                        return string.IsNullOrEmpty(k) ? r.Part.GetTitle() : k;
                    })
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Count();
                form.Text = rows.Count > 0 ? $"修改方程式（共 {partCount} 个零件）" : "修改方程式";
                form.Size = new Size(880, 520);
                form.StartPosition = FormStartPosition.CenterScreen;
                form.MinimumSize = new Size(640, 420);
                form.TopMost = true;

                var mainPanel = new TableLayoutPanel();
                mainPanel.Dock = DockStyle.Fill;
                mainPanel.ColumnCount = 1;
                mainPanel.RowCount = 5;
                mainPanel.Padding = new Padding(10);
                mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
                mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                var titleLabel = new Label
                {
                    Text = "选择一条方程式；同名变量请认准左侧 [零件: …] 标签。可多次「应用」后点「完成」。",
                    Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold, GraphicsUnit.Point),
                    AutoSize = true
                };
                mainPanel.Controls.Add(titleLabel, 0, 0);

                var listBox = new ListBox
                {
                    Dock = DockStyle.Fill,
                    IntegralHeight = false,
                    SelectionMode = SelectionMode.One,
                    Font = new Font("Consolas", 9F)
                };

                for (int i = 0; i < rows.Count; i++)
                {
                    var mgr = (EquationMgr)rows[i].Part.GetEquationMgr();
                    listBox.Items.Add(FormatListLine(rows[i], mgr));
                }

                mainPanel.Controls.Add(listBox, 0, 1);

                var inputPanel = new TableLayoutPanel();
                inputPanel.ColumnCount = 1;
                inputPanel.RowCount = 2;
                inputPanel.Dock = DockStyle.Fill;
                inputPanel.Padding = new Padding(0, 10, 0, 0);
                inputPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                inputPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                var valueLabel = new Label
                {
                    Text = "新值（支持表达式，如 150+231-18）:",
                    AutoSize = true,
                    Dock = DockStyle.Fill
                };
                inputPanel.Controls.Add(valueLabel, 0, 0);

                var textBox = new TextBox
                {
                    Dock = DockStyle.Top,
                    MinimumSize = new Size(300, 28),
                    Margin = new Padding(0, 6, 0, 0)
                };
                inputPanel.Controls.Add(textBox, 0, 1);

                var statusLabel = new Label
                {
                    Text = "",
                    AutoSize = true,
                    ForeColor = Color.DimGray
                };

                bool ActivatePartByRow(EquationListRow row)
                {
                    try
                    {
                        if (row?.Part == null)
                        {
                            return false;
                        }

                        row.Part.Visible = true;

                        string title = row.Part.GetTitle();
                        if (!string.IsNullOrEmpty(title))
                        {
                            int errors = 0;
                            swApp.ActivateDoc3(
                                title,
                                true,
                                (int)swRebuildOnActivation_e.swDontRebuildActiveDoc,
                                ref errors);
                        }

                        return row.Part.Visible;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"切换零件失败: {ex.Message}");
                        return false;
                    }
                }

                void RefreshListLine(int index)
                {
                    if (index < 0 || index >= rows.Count)
                    {
                        return;
                    }

                    var mgr = (EquationMgr)rows[index].Part.GetEquationMgr();
                    listBox.Items[index] = FormatListLine(rows[index], mgr);
                }

                void SyncTextFromSelection()
                {
                    if (listBox.SelectedIndex < 0)
                    {
                        return;
                    }

                    var row = rows[listBox.SelectedIndex];
                    if (!ActivatePartByRow(row))
                    {
                        statusLabel.Text = $"切换零件失败：{row.PartDisplayLabel}";
                        return;
                    }

                    var mgr = (EquationMgr)row.Part.GetEquationMgr();
                    if (mgr == null)
                    {
                        return;
                    }

                    string equation = mgr.get_Equation(row.EquationIndex);
                    textBox.Text = ExtractEquationValue(equation);
                    textBox.Focus();
                    textBox.SelectAll();
                    statusLabel.Text = $"已切换零件：{row.PartDisplayLabel}";
                }

                listBox.SelectedIndexChanged += (s, e) => SyncTextFromSelection();

                mainPanel.Controls.Add(inputPanel, 0, 2);
                mainPanel.Controls.Add(statusLabel, 0, 3);

                var buttonPanel = new FlowLayoutPanel
                {
                    FlowDirection = FlowDirection.RightToLeft,
                    AutoSize = true,
                    WrapContents = false,
                    Padding = new Padding(0, 10, 0, 0)
                };

                var cancelButton = new Button { Text = "取消", Width = 80, DialogResult = DialogResult.Cancel };
                var doneButton = new Button { Text = "完成", Width = 80, DialogResult = DialogResult.OK };
                var applyButton = new Button { Text = "应用", Width = 80 };

                buttonPanel.Controls.Add(cancelButton);
                buttonPanel.Controls.Add(doneButton);
                buttonPanel.Controls.Add(applyButton);

                mainPanel.Controls.Add(buttonPanel, 0, 4);

                form.Controls.Add(mainPanel);
                form.AcceptButton = applyButton;
                form.CancelButton = cancelButton;

                void ApplyCurrent()
                {
                    if (listBox.SelectedIndex < 0)
                    {
                        swApp.SendMsgToUser("请先选择一条方程式");
                        return;
                    }

                    string newValue = textBox.Text.Trim();
                    if (string.IsNullOrEmpty(newValue))
                    {
                        swApp.SendMsgToUser("请输入新的方程式值");
                        return;
                    }

                    var row = rows[listBox.SelectedIndex];
                    var mgr = (EquationMgr)row.Part.GetEquationMgr();
                    if (mgr == null)
                    {
                        return;
                    }

                    if (!ModifyEquationByIndex(mgr, row.Part, row.EquationIndex, newValue, swApp, row.PartDisplayLabel))
                    {
                        return;
                    }

                    try
                    {
                        row.Part.EditRebuild3();
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"重建失败: {ex.Message}");
                    }

                    RefreshListLine(listBox.SelectedIndex);
                    appliedFlag[0] = true;
                    statusLabel.Text = $"已应用：{DateTime.Now:HH:mm:ss}  [{row.PartDisplayLabel}] 第 {row.EquationIndex + 1} 条";
                }

                applyButton.Click += (s, e) => ApplyCurrent();

                form.Shown += (s, e) =>
                {
                    if (listBox.Items.Count > 0)
                    {
                        listBox.SelectedIndex = 0;
                    }
                    else
                    {
                        textBox.Focus();
                    }
                };

                return form.ShowDialog() == DialogResult.OK;
            }
        }

        /// <summary>
        /// 根据索引修改方程式
        /// </summary>
        private static bool ModifyEquationByIndex(
            EquationMgr eqnMgr,
            ModelDoc2 swModel,
            int selectedIndex,
            string newValue,
            SldWorks swApp,
            string? partDisplayLabel)
        {
            try
            {
                if (selectedIndex < 0 || selectedIndex >= eqnMgr.GetCount())
                {
                    Debug.WriteLine("选择的方程式索引无效");
                    swApp.SendMsgToUser("选择的方程式索引无效");
                    return false;
                }

                string calculatedValue = EvaluateExpression(newValue);

                string originalEquation = eqnMgr.get_Equation(selectedIndex);
                string[] originalParts = originalEquation.Split(new[] { '=' }, 2);

                if (originalParts.Length < 2)
                {
                    Debug.WriteLine("原始方程式格式错误");
                    swApp.SendMsgToUser("原始方程式格式错误");
                    return false;
                }

                string equationName = originalParts[0];

                string newEquation = $"{equationName}={calculatedValue}";

                eqnMgr.set_Equation(selectedIndex, newEquation);

                Debug.WriteLine($"方程式已更新: {newEquation}");

                var varDisplay = equationName.Trim().Replace("\r", " ").Replace("\n", " ");
                if (varDisplay.Length > 120)
                {
                    varDisplay = varDisplay.Substring(0, 120) + "…";
                }

                var valDisplay = (calculatedValue ?? "").Trim().Replace("\r", " ").Replace("\n", " ");
                if (valDisplay.Length > 80)
                {
                    valDisplay = valDisplay.Substring(0, 80) + "…";
                }

                string partTag = string.IsNullOrWhiteSpace(partDisplayLabel)
                    ? (swModel?.GetTitle() ?? "")
                    : partDisplayLabel.Trim();
                if (partTag.Length > 80)
                {
                    partTag = partTag.Substring(0, 80) + "…";
                }

                AiOperationBrief.Log($"改方程式：零件={partTag}，变量={varDisplay}，新值={valDisplay}");
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"修改方程式时出错: {ex.Message}");
                swApp?.SendMsgToUser($"修改方程式时出错: {ex.Message}");
                return false;
            }
        }

        private static string ExtractEquationValue(string equation)
        {
            if (string.IsNullOrWhiteSpace(equation))
            {
                return "";
            }

            var parts = equation.Split(new[] { '=' }, 2);
            if (parts.Length < 2)
            {
                return equation.Trim();
            }

            return parts[1].Trim();
        }

        private static string EvaluateExpression(string expression)
        {
            try
            {
                if (!expression.Contains('+') && !expression.Contains('-') &&
                    !expression.Contains('*') && !expression.Contains('/') &&
                    !expression.Contains('(') && !expression.Contains(')'))
                {
                    return expression;
                }

                var dataTable = new System.Data.DataTable();
                var result = dataTable.Compute(expression, "");

                if (result is double doubleResult)
                {
                    if (doubleResult == Math.Floor(doubleResult))
                    {
                        return ((int)doubleResult).ToString();
                    }

                    return doubleResult.ToString("F2");
                }

                if (result is int intResult)
                {
                    return intResult.ToString();
                }

                return result.ToString();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"表达式计算失败: {ex.Message}，使用原始值");
                return expression;
            }
        }
    }
}
