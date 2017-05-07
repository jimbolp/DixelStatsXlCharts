namespace DixelXlCharts
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.filePathTextBox = new System.Windows.Forms.TextBox();
            this.graphicsCheckBox = new System.Windows.Forms.CheckBox();
            this.printCheckBox = new System.Windows.Forms.CheckBox();
            this.startWorking = new System.Windows.Forms.Button();
            this.resultLabel = new System.Windows.Forms.Label();
            this.sheetNameLabel = new System.Windows.Forms.Label();
            this.debugTextBox = new System.Windows.Forms.RichTextBox();
            this.chartProgBar = new System.Windows.Forms.ProgressBar();
            this.convertProgBar = new System.Windows.Forms.ProgressBar();
            this.labelConverting = new System.Windows.Forms.Label();
            this.labelCharts = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // filePathTextBox
            // 
            this.filePathTextBox.AllowDrop = true;
            this.filePathTextBox.Location = new System.Drawing.Point(12, 32);
            this.filePathTextBox.Name = "filePathTextBox";
            this.filePathTextBox.Size = new System.Drawing.Size(210, 20);
            this.filePathTextBox.TabIndex = 1;
            this.filePathTextBox.DragDrop += new System.Windows.Forms.DragEventHandler(this.FilePathTextBox_DragDrop);
            this.filePathTextBox.DragOver += new System.Windows.Forms.DragEventHandler(this.FilePathTextBox_DragOver);
            // 
            // graphicsCheckBox
            // 
            this.graphicsCheckBox.AutoSize = true;
            this.graphicsCheckBox.Location = new System.Drawing.Point(12, 58);
            this.graphicsCheckBox.Name = "graphicsCheckBox";
            this.graphicsCheckBox.Size = new System.Drawing.Size(143, 17);
            this.graphicsCheckBox.TabIndex = 8;
            this.graphicsCheckBox.Text = "Създаване на графики";
            this.graphicsCheckBox.UseVisualStyleBackColor = true;
            this.graphicsCheckBox.CheckedChanged += new System.EventHandler(this.GraphicsCheckBox_CheckedChanged);
            // 
            // printCheckBox
            // 
            this.printCheckBox.AutoSize = true;
            this.printCheckBox.Enabled = false;
            this.printCheckBox.Location = new System.Drawing.Point(12, 81);
            this.printCheckBox.Name = "printCheckBox";
            this.printCheckBox.Size = new System.Drawing.Size(138, 17);
            this.printCheckBox.TabIndex = 7;
            this.printCheckBox.Text = "Принтирай графиките";
            this.printCheckBox.UseVisualStyleBackColor = true;
            // 
            // startWorking
            // 
            this.startWorking.Location = new System.Drawing.Point(275, 32);
            this.startWorking.Name = "startWorking";
            this.startWorking.Size = new System.Drawing.Size(76, 66);
            this.startWorking.TabIndex = 9;
            this.startWorking.Text = "Start";
            this.startWorking.UseVisualStyleBackColor = true;
            this.startWorking.Click += new System.EventHandler(this.StartWorking_Click);
            // 
            // resultLabel
            // 
            this.resultLabel.AutoEllipsis = true;
            this.resultLabel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.resultLabel.Location = new System.Drawing.Point(0, 355);
            this.resultLabel.Name = "resultLabel";
            this.resultLabel.Size = new System.Drawing.Size(326, 132);
            this.resultLabel.TabIndex = 10;
            // 
            // sheetNameLabel
            // 
            this.sheetNameLabel.AutoSize = true;
            this.sheetNameLabel.Location = new System.Drawing.Point(9, 146);
            this.sheetNameLabel.Name = "sheetNameLabel";
            this.sheetNameLabel.Size = new System.Drawing.Size(0, 13);
            this.sheetNameLabel.TabIndex = 11;
            // 
            // debugTextBox
            // 
            this.debugTextBox.Location = new System.Drawing.Point(533, 311);
            this.debugTextBox.Name = "debugTextBox";
            this.debugTextBox.Size = new System.Drawing.Size(177, 173);
            this.debugTextBox.TabIndex = 13;
            this.debugTextBox.Text = "";
            // 
            // chartProgBar
            // 
            this.chartProgBar.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.chartProgBar.Location = new System.Drawing.Point(0, 167);
            this.chartProgBar.Name = "chartProgBar";
            this.chartProgBar.Size = new System.Drawing.Size(383, 23);
            this.chartProgBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.chartProgBar.TabIndex = 14;
            // 
            // convertProgBar
            // 
            this.convertProgBar.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.convertProgBar.Location = new System.Drawing.Point(0, 144);
            this.convertProgBar.Name = "convertProgBar";
            this.convertProgBar.Size = new System.Drawing.Size(383, 23);
            this.convertProgBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.convertProgBar.TabIndex = 15;
            // 
            // labelConverting
            // 
            this.labelConverting.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.labelConverting.AutoSize = true;
            this.labelConverting.BackColor = System.Drawing.Color.Transparent;
            this.labelConverting.Enabled = false;
            this.labelConverting.Location = new System.Drawing.Point(5, 149);
            this.labelConverting.Name = "labelConverting";
            this.labelConverting.Size = new System.Drawing.Size(54, 13);
            this.labelConverting.TabIndex = 16;
            this.labelConverting.Text = "Loading...";
            this.labelConverting.Visible = false;
            // 
            // labelCharts
            // 
            this.labelCharts.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.labelCharts.AutoSize = true;
            this.labelCharts.BackColor = System.Drawing.Color.Transparent;
            this.labelCharts.Enabled = false;
            this.labelCharts.Location = new System.Drawing.Point(5, 171);
            this.labelCharts.Name = "labelCharts";
            this.labelCharts.Size = new System.Drawing.Size(54, 13);
            this.labelCharts.TabIndex = 17;
            this.labelCharts.Text = "Loading...";
            this.labelCharts.Visible = false;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(383, 190);
            this.Controls.Add(this.labelCharts);
            this.Controls.Add(this.labelConverting);
            this.Controls.Add(this.convertProgBar);
            this.Controls.Add(this.chartProgBar);
            this.Controls.Add(this.debugTextBox);
            this.Controls.Add(this.sheetNameLabel);
            this.Controls.Add(this.resultLabel);
            this.Controls.Add(this.startWorking);
            this.Controls.Add(this.graphicsCheckBox);
            this.Controls.Add(this.printCheckBox);
            this.Controls.Add(this.filePathTextBox);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainForm";
            this.Text = "Създаване на графики";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox filePathTextBox;
        private System.Windows.Forms.CheckBox graphicsCheckBox;
        private System.Windows.Forms.CheckBox printCheckBox;
        private System.Windows.Forms.Button startWorking;
        private System.Windows.Forms.Label resultLabel;
        private System.Windows.Forms.Label sheetNameLabel;
        private System.Windows.Forms.RichTextBox debugTextBox;
        private System.Windows.Forms.ProgressBar chartProgBar;
        private System.Windows.Forms.ProgressBar convertProgBar;
        private System.Windows.Forms.Label labelConverting;
        private System.Windows.Forms.Label labelCharts;
    }
}

