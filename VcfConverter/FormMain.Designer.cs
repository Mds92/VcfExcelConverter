
namespace VcfConverter
{
    partial class FormMain
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormMain));
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxSrcFilePath = new System.Windows.Forms.TextBox();
            this.buttonSelectVcfFile = new System.Windows.Forms.Button();
            this.buttonStartConverting = new System.Windows.Forms.Button();
            this.textBoxLog = new System.Windows.Forms.TextBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "contacts.vcf";
            this.openFileDialog1.Filter = "All Files|*.vcf;*.xlsx";
            this.openFileDialog1.Title = "Select VCF File";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(94, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "Source File Path:";
            // 
            // textBoxSrcFilePath
            // 
            this.textBoxSrcFilePath.Location = new System.Drawing.Point(112, 6);
            this.textBoxSrcFilePath.Name = "textBoxSrcFilePath";
            this.textBoxSrcFilePath.Size = new System.Drawing.Size(428, 23);
            this.textBoxSrcFilePath.TabIndex = 1;
            this.textBoxSrcFilePath.TextChanged += new System.EventHandler(this.textBoxSrcFilePath_TextChanged);
            // 
            // buttonSelectVcfFile
            // 
            this.buttonSelectVcfFile.Location = new System.Drawing.Point(546, 6);
            this.buttonSelectVcfFile.Name = "buttonSelectVcfFile";
            this.buttonSelectVcfFile.Size = new System.Drawing.Size(35, 23);
            this.buttonSelectVcfFile.TabIndex = 2;
            this.buttonSelectVcfFile.Text = "...";
            this.buttonSelectVcfFile.UseVisualStyleBackColor = true;
            this.buttonSelectVcfFile.Click += new System.EventHandler(this.buttonSelectVcfFile_Click);
            // 
            // buttonStartConverting
            // 
            this.buttonStartConverting.Location = new System.Drawing.Point(44, 35);
            this.buttonStartConverting.Name = "buttonStartConverting";
            this.buttonStartConverting.Size = new System.Drawing.Size(132, 64);
            this.buttonStartConverting.TabIndex = 4;
            this.buttonStartConverting.Text = "Convert";
            this.buttonStartConverting.UseVisualStyleBackColor = true;
            this.buttonStartConverting.Click += new System.EventHandler(this.buttonStartConverting_Click);
            // 
            // textBoxLog
            // 
            this.textBoxLog.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.textBoxLog.Location = new System.Drawing.Point(0, 105);
            this.textBoxLog.Multiline = true;
            this.textBoxLog.Name = "textBoxLog";
            this.textBoxLog.ReadOnly = true;
            this.textBoxLog.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBoxLog.Size = new System.Drawing.Size(584, 106);
            this.textBoxLog.TabIndex = 5;
            this.textBoxLog.WordWrap = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(221, 35);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(319, 64);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBox1.TabIndex = 7;
            this.pictureBox1.TabStop = false;
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(584, 211);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.textBoxLog);
            this.Controls.Add(this.buttonStartConverting);
            this.Controls.Add(this.buttonSelectVcfFile);
            this.Controls.Add(this.textBoxSrcFilePath);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "FormMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "VCF Excel Converter";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxSrcFilePath;
        private System.Windows.Forms.Button buttonSelectVcfFile;
        private System.Windows.Forms.Button buttonStartConverting;
        private System.Windows.Forms.TextBox textBoxLog;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}

