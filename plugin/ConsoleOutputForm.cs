using System;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace SolidWorksAddinStudy
{
    public class ConsoleOutputForm : Form
    {
        private TextBox outputTextBox;
        private StringBuilder buffer = new StringBuilder();
        private TextWriter? originalOut;
        
        public ConsoleOutputForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "实时输出";
            this.Size = new Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;

            outputTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                WordWrap = false,
                ReadOnly = true,
                Font = new Font("Consolas", 9F),
                BackColor = Color.White,
                ForeColor = Color.Black
            };

            this.Controls.Add(outputTextBox);
        }

        public void StartIntercept()
        {
            originalOut = Console.Out;
            Console.SetOut(new ConsoleOutputWriter(this));
        }

        public void StopIntercept()
        {
            if (originalOut != null)
            {
                Console.SetOut(originalOut);
            }
        }

        public void AppendText(string text)
        {
            if (outputTextBox.InvokeRequired)
            {
                outputTextBox.Invoke(new Action<string>(AppendText), text);
                return;
            }

            buffer.Append(text);
            outputTextBox.Text = buffer.ToString();
            outputTextBox.SelectionStart = outputTextBox.Text.Length;
            outputTextBox.ScrollToCaret();
        }

        private class ConsoleOutputWriter : TextWriter
        {
            private ConsoleOutputForm parent;
            public ConsoleOutputWriter(ConsoleOutputForm parent) => this.parent = parent;
            public override Encoding Encoding => Encoding.UTF8;
            public override void Write(char value) => parent.AppendText(value.ToString());
            public override void Write(string? value) => parent.AppendText(value ?? "");
            public override void WriteLine(string? value) => parent.AppendText((value ?? "") + Environment.NewLine);
        }
    }
}
