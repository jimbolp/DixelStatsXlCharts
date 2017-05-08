using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace DixelXlCharts
{
    public partial class MainForm : Form
    {
        private static MainForm form = null;
        private bool isProcessRunning = false;
        private delegate void EnableDelegateSave(string text);
        private delegate void EnableDelegateProgBar(int val, bool max);

        public static bool TempCharts { get; set; }
        public static bool HumidCharts { get; set; }
        public static string SaveFilePath { get; set; } = null;
        private string loadedFile = "";
        public MainForm()
        {
            InitializeComponent();
            graphicsCheckBox.Checked = true;
            TempCharts = tempChckBox.Checked;
            HumidCharts = humidChckBox.Checked;
            this.SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
            convertProgBar.CreateGraphics().TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
            chartProgBar.CreateGraphics().TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
            this.UpdateStyles();
            form = this;
            Icon = Resources.icons._002_analytics_1;
        }
        public static void ProgressBar(int val, bool max)
        {
            form?.ProgBar(val, max);
        }

        public static void ConvProgBar(int val, bool max)
        {
            form?.CProgBar(val, max);
        }

        private void CProgBar(int val, bool max)
        {
            if (InvokeRequired)
            {
                Invoke(new EnableDelegateProgBar(CProgBar), new object[] { val, max });
                return;
            }

            switch (max)
            {
                case true:
                    if (convertProgBar.Maximum < val || val == 0)
                    {
                        convertProgBar.Maximum = val;
                    }
                    break;
                case false:
                    if (convertProgBar.Value < val && val < convertProgBar.Maximum)
                    {
                        if (convertProgBar.Maximum != 0)
                        {
                            convertProgBar.Refresh();
                            int percent = (int)(((double)convertProgBar.Value / (double)convertProgBar.Maximum) * 100);
                            convertProgBar.CreateGraphics().DrawString(percent.ToString() + "% Converting...", new Font("Arial", (float)9.00, FontStyle.Regular), Brushes.Black, new PointF(convertProgBar.Width / 2 - 35, convertProgBar.Height / 2 - 7));
                            
                        }
                        convertProgBar.Value = val;
                    }
                    break;
            }
        }
        private void ProgBar(int val, bool max)
        {
            if (InvokeRequired)
            {
                Invoke(new EnableDelegateProgBar(ProgBar), new object[] { val, max });
                return;
            }
            
            switch (max)
            {
                case true:
                    if(chartProgBar.Maximum < val || val == 0)
                    {
                        chartProgBar.Maximum = val;
                    }
                    break;
                case false:
                    if (chartProgBar.Value < val && val < chartProgBar.Maximum)
                    {
                        if (chartProgBar.Maximum != 0)
                        {
                            chartProgBar.Refresh();
                            int percent = (int)(((double)chartProgBar.Value / (double)chartProgBar.Maximum) * 100);
                            chartProgBar.CreateGraphics().DrawString(percent.ToString() + "% Creating Charts...", new Font("Arial", (float)9.00, FontStyle.Regular), Brushes.Black, new PointF(chartProgBar.Width / 2 - 35, chartProgBar.Height / 2 - 7));
                            //labelCharts.Text = (int)(((double)chartProgBar.Value / (double)chartProgBar.Maximum )* 100) + "% Creating Charts...";
                        }
                        else
                        {
                            //labelCharts.Text = "Loading...";
                        }
                        chartProgBar.Value = val;
                    }
                    break;
            }
        }
        /*public static void WriteIntoLabel(string text, int l)
        {
                if (form != null)
                    form.WriteText(text, l);
        }
        private void WriteText(string text, int l)
        {
            // If this returns true, it means it was called from an external thread.
            if (InvokeRequired)
            {
                // Create a delegate of this method and let the form run it.
                this.Invoke(new EnableDelegate(WriteText), new object[] { text, l });
                return;
            }

            // Set TextBox or RichTextBox
            switch (l)
            {
                case 1:
                    resultLabel.Text = text;
                    break;
                case 2:
                    debugTextBox.Text = text;
                    //debugTextBox.Text += text + Environment.NewLine;
                    break;
            }
        }//*/
        private void FilePathTextBox_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
        }

        private void FilePathTextBox_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files != null && files.Length != 0)
            {
                filePathTextBox.Text = files[0];
                loadedFile = filePathTextBox.Text;
            }
        }

        private void StartWorking_Click(object sender, EventArgs e)
        {
            if (isProcessRunning)
            {
                MessageBox.Show("The process is already running!!");
                return;
            }
            if (!graphicsCheckBox.Checked)
            {
                DialogResult dr =
                    MessageBox.Show(
                        "Не сте избрали опцията за създаване на графики! Сигурни ли сте, че искате да се създадат графики?",
                        "Внимание!", MessageBoxButtons.YesNo);
                if (dr == DialogResult.No)
                {
                    return;
                }
                graphicsCheckBox.Checked = true;
            }
            else
            {
                if (!tempChckBox.Checked && !humidChckBox.Checked)
                {
                    MessageBox.Show("Изберете \"Температура\", \"Влажност\" или и двете!");
                    return;
                }
            }
            DixelData dxData = null;
            try
            {
                Thread thr1 = new Thread(() =>
                {
                    isProcessRunning = true;
                    Thread.CurrentThread.IsBackground = false;
                    try
                    {
                        dxData = new DixelData(filePathTextBox.Text, printCheckBox.Checked);
                        dxData.LoadData();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        isProcessRunning = false;
                        return;
                    }
                    dxData.SaveAndClose();
                    isProcessRunning = false;
                });
                thr1.Start();
            }
            catch (COMException)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            dxData?.Dispose();
            
            isProcessRunning = false;
        }
        public static string SaveDialogBox(string saveFileDir)
        {
            if(form != null)
            {
                form.SaveBox(saveFileDir);
                return SaveFilePath;
            }
            return null;
        }
        private void SaveBox(string saveFileDir)
        {
            if (InvokeRequired)
            {
                // Create a delegate of this method and let the form run it.
                this.Invoke(new EnableDelegateSave(SaveBox), new object[] { saveFileDir });
                return;
            }
            SaveFileDialog saveFileDialog;
            if (Path.GetExtension(loadedFile) == ".xls")
            {
                saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel 97-2003 Workbook|*.xls",
                    Title = "Save As",
                    DefaultExt = ".xls",
                    InitialDirectory = saveFileDir
                };
            }
            else
            {
                saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Workbook|*.xlsx; *.xlsm",
                    Title = "Save As",
                    DefaultExt = ".xlsx",
                    InitialDirectory = saveFileDir
                };
            }
            saveFileDialog.AddExtension = true;

            DialogResult dr = saveFileDialog.ShowDialog();
            if (dr == DialogResult.OK && !string.IsNullOrEmpty(saveFileDialog.FileName))
            {
                SaveFilePath = saveFileDialog.FileName;
            }
            else
            {
                SaveFilePath = null;
            }
        }
        private void GraphicsCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (graphicsCheckBox.Checked)
            {
                printCheckBox.Enabled = true;
                tempChckBox.Enabled = true;
                tempChckBox.Checked = true;
                humidChckBox.Enabled = true;
                humidChckBox.Checked = true;
            }
            else
            {
                printCheckBox.Checked = false;
                printCheckBox.Enabled = false;
                tempChckBox.Checked = false;
                humidChckBox.Checked = false;
                tempChckBox.Enabled = false;
                humidChckBox.Enabled = false;
            }
        }

        private void tempChckBox_CheckedChanged(object sender, EventArgs e)
        {
            TempCharts = tempChckBox.Checked;
        }

        private void humidChckBox_CheckedChanged(object sender, EventArgs e)
        {
            HumidCharts = humidChckBox.Checked;
        }
    }
}
