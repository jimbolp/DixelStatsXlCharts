using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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
        private delegate void EnableDelegate(string text, int l);

        private delegate void EnableDelegateSave(string text);
        public static string SaveFilePath { get; set; } = null;
        public MainForm()
        {
            InitializeComponent();
            form = this;
        }
        public static void WriteIntoLabel(string text, int l)
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
                return; // Important
            }

            // Set textBox
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
            }
        }

        private void StartWorking_Click(object sender, EventArgs e)
        {

            DixelData dxData = null;
            try
            {
                Thread thr1 = new Thread(() =>
                {
                    Thread.CurrentThread.IsBackground = false;
                    dxData = new DixelData(filePathTextBox.Text, printCheckBox.Checked);
                    dxData.LoadData();
                    dxData.SaveAndClose();
                });
                thr1.Start();
            }
            catch (COMException)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            if(dxData != null)
            {
                dxData.Dispose();
            }
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
                return; // Important
            }
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Workbook|*.xlsx; *.xlsm|Excel 97-2003 Workbook|*.xls",
                Title = "Save As",
                DefaultExt = "xlsx",
                InitialDirectory = saveFileDir
            };
            saveFileDialog.AddExtension = true;

            DialogResult dr = saveFileDialog.ShowDialog();
            if (dr == DialogResult.OK)
            {
                SaveFilePath = saveFileDialog.FileName;
            }
        }
        private void GraphicsCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (graphicsCheckBox.Checked)
            {
                printCheckBox.Enabled = true;
            }
            else
            {
                printCheckBox.Checked = false;
                printCheckBox.Enabled = false;
            }
        }
    }
}
