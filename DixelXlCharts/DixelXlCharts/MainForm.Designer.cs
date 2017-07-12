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
            this.components = new System.ComponentModel.Container();
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
            this.tempChckBox = new System.Windows.Forms.CheckBox();
            this.humidChckBox = new System.Windows.Forms.CheckBox();
            this.specialChckBox = new System.Windows.Forms.CheckBox();
            this.specialTip = new System.Windows.Forms.ToolTip(this.components);
            this.browseFileBtn = new System.Windows.Forms.Button();
            this.printExistingBtn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // filePathTextBox
            // 
            this.filePathTextBox.AllowDrop = true;
            this.filePathTextBox.Location = new System.Drawing.Point(12, 32);
            this.filePathTextBox.Name = "filePathTextBox";
            this.filePathTextBox.Size = new System.Drawing.Size(232, 20);
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
            this.printCheckBox.Location = new System.Drawing.Point(12, 143);
            this.printCheckBox.Name = "printCheckBox";
            this.printCheckBox.Size = new System.Drawing.Size(138, 17);
            this.printCheckBox.TabIndex = 7;
            this.printCheckBox.Text = "Принтирай графиките";
            this.printCheckBox.UseVisualStyleBackColor = true;
            // 
            // startWorking
            // 
            this.startWorking.BackColor = System.Drawing.Color.Transparent;
            this.startWorking.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.startWorking.Image = ((System.Drawing.Image)(resources.GetObject("startWorking.Image")));
            this.startWorking.Location = new System.Drawing.Point(275, 32);
            this.startWorking.Name = "startWorking";
            this.startWorking.Size = new System.Drawing.Size(76, 66);
            this.startWorking.TabIndex = 9;
            this.startWorking.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.startWorking.UseVisualStyleBackColor = false;
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
            this.chartProgBar.Location = new System.Drawing.Point(0, 208);
            this.chartProgBar.Name = "chartProgBar";
            this.chartProgBar.Size = new System.Drawing.Size(383, 23);
            this.chartProgBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.chartProgBar.TabIndex = 14;
            // 
            // convertProgBar
            // 
            this.convertProgBar.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.convertProgBar.Location = new System.Drawing.Point(0, 185);
            this.convertProgBar.Name = "convertProgBar";
            this.convertProgBar.Size = new System.Drawing.Size(383, 23);
            this.convertProgBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.convertProgBar.TabIndex = 15;
            // 
            // tempChckBox
            // 
            this.tempChckBox.AutoSize = true;
            this.tempChckBox.Enabled = false;
            this.tempChckBox.Location = new System.Drawing.Point(21, 81);
            this.tempChckBox.Name = "tempChckBox";
            this.tempChckBox.Size = new System.Drawing.Size(93, 17);
            this.tempChckBox.TabIndex = 18;
            this.tempChckBox.Text = "Температура";
            this.tempChckBox.UseVisualStyleBackColor = true;
            this.tempChckBox.CheckedChanged += new System.EventHandler(this.tempChckBox_CheckedChanged);
            // 
            // humidChckBox
            // 
            this.humidChckBox.AutoSize = true;
            this.humidChckBox.Enabled = false;
            this.humidChckBox.Location = new System.Drawing.Point(21, 97);
            this.humidChckBox.Name = "humidChckBox";
            this.humidChckBox.Size = new System.Drawing.Size(76, 17);
            this.humidChckBox.TabIndex = 19;
            this.humidChckBox.Text = "Влажност";
            this.humidChckBox.UseVisualStyleBackColor = true;
            this.humidChckBox.CheckedChanged += new System.EventHandler(this.humidChckBox_CheckedChanged);
            // 
            // specialChckBox
            // 
            this.specialChckBox.AutoSize = true;
            this.specialChckBox.Location = new System.Drawing.Point(38, 120);
            this.specialChckBox.Name = "specialChckBox";
            this.specialChckBox.Size = new System.Drawing.Size(118, 17);
            this.specialChckBox.TabIndex = 20;
            this.specialChckBox.Text = "Специален случай";
            this.specialTip.SetToolTip(this.specialChckBox, "Специален случай е когато в подадения файл,\r\nчетенията са само за влажност\r\nи сто" +
        "йностите са във втората колона на Excel файла.\r\nВъв всички останали случай,\r\nтаз" +
        "и опция ще изкара грешни графики!");
            this.specialChckBox.UseVisualStyleBackColor = true;
            this.specialChckBox.CheckedChanged += new System.EventHandler(this.specialChckBox_CheckedChanged);
            // 
            // specialTip
            // 
            this.specialTip.AutomaticDelay = 0;
            this.specialTip.AutoPopDelay = 0;
            this.specialTip.InitialDelay = 500;
            this.specialTip.IsBalloon = true;
            this.specialTip.OwnerDraw = true;
            this.specialTip.ReshowDelay = 152;
            this.specialTip.ShowAlways = true;
            this.specialTip.ToolTipIcon = System.Windows.Forms.ToolTipIcon.Warning;
            this.specialTip.ToolTipTitle = "Внимание!";
            // 
            // browseFileBtn
            // 
            this.browseFileBtn.Location = new System.Drawing.Point(275, 104);
            this.browseFileBtn.Name = "browseFileBtn";
            this.browseFileBtn.Size = new System.Drawing.Size(76, 44);
            this.browseFileBtn.TabIndex = 21;
            this.browseFileBtn.Text = "Зареди файл";
            this.browseFileBtn.UseVisualStyleBackColor = true;
            this.browseFileBtn.Click += new System.EventHandler(this.browseFileBtn_Click);
            // 
            // printExistingBtn
            // 
            this.printExistingBtn.Location = new System.Drawing.Point(275, 154);
            this.printExistingBtn.Name = "printExistingBtn";
            this.printExistingBtn.Size = new System.Drawing.Size(75, 23);
            this.printExistingBtn.TabIndex = 22;
            this.printExistingBtn.Text = "button1";
            this.printExistingBtn.UseVisualStyleBackColor = true;
            this.printExistingBtn.Visible = false;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(383, 231);
            this.Controls.Add(this.printExistingBtn);
            this.Controls.Add(this.browseFileBtn);
            this.Controls.Add(this.specialChckBox);
            this.Controls.Add(this.humidChckBox);
            this.Controls.Add(this.tempChckBox);
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
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
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
        private System.Windows.Forms.CheckBox tempChckBox;
        private System.Windows.Forms.CheckBox humidChckBox;
        private System.Windows.Forms.CheckBox specialChckBox;
        private System.Windows.Forms.ToolTip specialTip;
        private System.Windows.Forms.Button browseFileBtn;
        private System.Windows.Forms.Button printExistingBtn;
    }
}

