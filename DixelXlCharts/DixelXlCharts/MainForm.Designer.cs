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
            this.filePathTextBox = new System.Windows.Forms.TextBox();
            this.graphicsCheckBox = new System.Windows.Forms.CheckBox();
            this.printCheckBox = new System.Windows.Forms.CheckBox();
            this.startWorking = new System.Windows.Forms.Button();
            this.resultLabel = new System.Windows.Forms.Label();
            this.sheetNameLabel = new System.Windows.Forms.Label();
            this.debugTextBox = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // filePathTextBox
            // 
            this.filePathTextBox.AllowDrop = true;
            this.filePathTextBox.Location = new System.Drawing.Point(12, 32);
            this.filePathTextBox.Name = "filePathTextBox";
            this.filePathTextBox.Size = new System.Drawing.Size(210, 20);
            this.filePathTextBox.TabIndex = 1;
            this.filePathTextBox.DragDrop += new System.Windows.Forms.DragEventHandler(this.filePathTextBox_DragDrop);
            this.filePathTextBox.DragOver += new System.Windows.Forms.DragEventHandler(this.filePathTextBox_DragOver);
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
            this.graphicsCheckBox.CheckedChanged += new System.EventHandler(this.graphicsCheckBox_CheckedChanged);
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
            this.startWorking.Click += new System.EventHandler(this.startWorking_Click);
            // 
            // resultLabel
            // 
            this.resultLabel.AutoEllipsis = true;
            this.resultLabel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.resultLabel.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.resultLabel.Location = new System.Drawing.Point(0, 159);
            this.resultLabel.Name = "resultLabel";
            this.resultLabel.Size = new System.Drawing.Size(572, 132);
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
            this.debugTextBox.Dock = System.Windows.Forms.DockStyle.Right;
            this.debugTextBox.Location = new System.Drawing.Point(395, 0);
            this.debugTextBox.Name = "debugTextBox";
            this.debugTextBox.Size = new System.Drawing.Size(177, 159);
            this.debugTextBox.TabIndex = 13;
            this.debugTextBox.Text = "";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(572, 291);
            this.Controls.Add(this.debugTextBox);
            this.Controls.Add(this.sheetNameLabel);
            this.Controls.Add(this.resultLabel);
            this.Controls.Add(this.startWorking);
            this.Controls.Add(this.graphicsCheckBox);
            this.Controls.Add(this.printCheckBox);
            this.Controls.Add(this.filePathTextBox);
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
    }
}

