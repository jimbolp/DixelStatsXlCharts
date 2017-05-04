using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DixelXlCharts
{   
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void filePathTextBox_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
        }

        private void filePathTextBox_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files != null && files.Length != 0)
            {
                filePathTextBox.Text = files[0];
            }
        }

        private void startWorking_Click(object sender, EventArgs e)
        {
            DixelData dxData = null;
            try
            {
                dxData = new DixelData(filePathTextBox.Text, printCheckBox.Checked);
                dxData.LoadData();
                dxData.SaveAndClose();
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

        private void graphicsCheckBox_CheckedChanged(object sender, EventArgs e)
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
