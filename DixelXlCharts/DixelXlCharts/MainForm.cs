using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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
        private delegate void EnableDelegateConvProgBar(int val, bool max, int number, int count);
        public static bool SpecialCase { get; set; }
        public static bool TempCharts { get; set; }
        public static bool HumidCharts { get; set; }
        public static string SaveFilePath { get; set; } = null;
        public static bool isCancellationRequested { get; set; } = false;
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
            specialTip.SetToolTip(this.specialChckBox, "Специален случай е когато в подадения файл, четенията са само за влажност и стойностите са във втората колона на Excel файла."
                + Environment.NewLine +
                "Във всички останали случай, тази опция ще изкара грешни графики!");
        }
        

        public static void ProgressBar(int val, bool max)
        {
            form?.ProgBar(val, max);
        }

        public static void ConvProgBar(int val, bool max, int number, int count)
        {
            form?.CProgBar(val, max, number, count);
        }

        private void CProgBar(int val, bool max, int number, int count)
        {
            if (InvokeRequired)
            {
                try
                {
                    Invoke(new EnableDelegateConvProgBar(CProgBar), new object[] { val, max, number, count });
                    return;
                }
                catch (Exception)
                {
                    return;
                }
            }
           lock(convertProgBar)
            {
                switch (max)
                {
                    case true:
                        if (convertProgBar.Maximum < val || val == 1 || val == 0)
                        {
                            convertProgBar.Maximum = val;
                        }
                        break;
                    case false:
                        if (convertProgBar.Value < val && val <= convertProgBar.Maximum)
                        {
                            if (convertProgBar.Maximum != 0)
                            {
                                convertProgBar.Refresh();
                                int percent = (int)(((double)convertProgBar.Value / (double)convertProgBar.Maximum) * 100);
                                using (Graphics gr = convertProgBar.CreateGraphics())
                                {
                                    gr.DrawString(percent.ToString() + "% Converting... Page "+number + "/" + count, new Font("Arial", (float)9.00, FontStyle.Regular), Brushes.Black, new PointF(convertProgBar.Width / 2 - 60, convertProgBar.Height / 2 - 7));
                                }
                            }
                            convertProgBar.Value = val;
                        }
                        if(val == 0)
                            convertProgBar.Value = val;
                        break;
                }
            }
        }
        private void ProgBar(int val, bool max)
        {
            if (InvokeRequired)
            {
                Invoke(new EnableDelegateProgBar(ProgBar), new object[] { val, max });
                return;
            }
            lock (chartProgBar)
            {
                switch (max)
                {
                    case true:
                        if (chartProgBar.Maximum < val || val == 0)
                        {
                            chartProgBar.Maximum = val;
                        }
                        break;
                    case false:
                        if (chartProgBar.Value < val && val <= chartProgBar.Maximum)
                        {
                            if (chartProgBar.Maximum != 0)
                            {
                                chartProgBar.Refresh();
                                int percent = (int)(((double)chartProgBar.Value / (double)chartProgBar.Maximum) * 100);
                                using (Graphics gr = chartProgBar.CreateGraphics())
                                {
                                    gr.DrawString(percent.ToString() + "% Creating Charts...", new Font("Arial", (float)9.00, FontStyle.Regular), Brushes.Black, new PointF(chartProgBar.Width / 2 - 35, chartProgBar.Height / 2 - 7));
                                }
                            }
                            chartProgBar.Value = val;
                        }
                        if(val == 0)
                            chartProgBar.Value = val;
                        break;
                }
            }
        }
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
            convertProgBar.Value = 0;
            chartProgBar.Value = 0;
            convertProgBar.Maximum = 1;
            chartProgBar.Maximum = 1;
            if (isProcessRunning)
            {
                MessageBox.Show("The process is already running!!");
                return;
            }
            if (!graphicsCheckBox.Checked)
            {
                DialogResult dr = MessageBox.Show(
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
                    //Thread.CurrentThread.IsBackground = false;
                    try
                    {
                        dxData = new DixelData(filePathTextBox.Text, printCheckBox.Checked);
                        dxData.LoadData();
                        dxData.SaveAndClose();//*/
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        if(dxData != null)
                            dxData.Dispose();
                        isProcessRunning = false;
                        
                        return;
                    }
                    isProcessRunning = false;
                });
                thr1.Start();
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
                    InitialDirectory = saveFileDir ?? Directory.GetCurrentDirectory()
                };
            }
            else
            {
                saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Workbook|*.xlsx; *.xlsm",
                    Title = "Save As",
                    DefaultExt = ".xlsx",
                    InitialDirectory = saveFileDir ?? Directory.GetCurrentDirectory()
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
            if(!tempChckBox.Checked && !humidChckBox.Checked)
            {
                printCheckBox.Checked = false;
                printCheckBox.Enabled = false;
            }
            if (!tempChckBox.Checked && humidChckBox.Checked)
            {
                specialChckBox.Enabled = true;
            }
            else if (tempChckBox.Checked)
            {
                printCheckBox.Enabled = true;
                specialChckBox.Checked = false;
                specialChckBox.Enabled = false;
            }
        }

        private void humidChckBox_CheckedChanged(object sender, EventArgs e)
        {
            HumidCharts = humidChckBox.Checked;
            if (!humidChckBox.Checked && !tempChckBox.Checked)
            {
                printCheckBox.Checked = false;
                printCheckBox.Enabled = false;
            }
            if (humidChckBox.Checked)
            {
                printCheckBox.Enabled = true;
                if (!tempChckBox.Checked)
                {
                    specialChckBox.Enabled = true;
                }
                else
                {
                    specialChckBox.Checked = false;
                    specialChckBox.Enabled = false;
                }
            }
            else
            {
                specialChckBox.Checked = false;
                specialChckBox.Enabled = false;
            }
        }

        private void specialChckBox_CheckedChanged(object sender, EventArgs e)
        {
            SpecialCase = specialChckBox.Checked;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            //isCancellationRequested = true;
            if (Marshal.AreComObjectsAvailableForCleanup())
            {
                Marshal.CleanupUnusedObjectsInCurrentContext();
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void browseFileBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = false;            
            if(DialogResult.OK == ofd.ShowDialog())
            {
                filePathTextBox.Text = ofd.FileName;
            }
        }
    }
}
