namespace CombileFileExcel
{
    partial class Form1
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
            this.btnChoiceMutipleFile = new System.Windows.Forms.Button();
            this.btnChooseSaveFile = new System.Windows.Forms.Button();
            this.txtFileName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnCreateFileExcel = new System.Windows.Forms.Button();
            this.btnChooseFolder = new System.Windows.Forms.Button();
            this.btnReChooseFile = new System.Windows.Forms.Button();
            this.btnCopyFilter = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnChoiceMutipleFile
            // 
            this.btnChoiceMutipleFile.Location = new System.Drawing.Point(73, 177);
            this.btnChoiceMutipleFile.Name = "btnChoiceMutipleFile";
            this.btnChoiceMutipleFile.Size = new System.Drawing.Size(75, 23);
            this.btnChoiceMutipleFile.TabIndex = 0;
            this.btnChoiceMutipleFile.Text = "Chọn file";
            this.btnChoiceMutipleFile.UseVisualStyleBackColor = true;
            this.btnChoiceMutipleFile.Click += new System.EventHandler(this.btnChoiceMutipleFile_Click);
            // 
            // btnChooseSaveFile
            // 
            this.btnChooseSaveFile.Location = new System.Drawing.Point(561, 68);
            this.btnChooseSaveFile.Name = "btnChooseSaveFile";
            this.btnChooseSaveFile.Size = new System.Drawing.Size(117, 23);
            this.btnChooseSaveFile.TabIndex = 2;
            this.btnChooseSaveFile.Text = "Chọn nơi lưu file";
            this.btnChooseSaveFile.UseVisualStyleBackColor = true;
            this.btnChooseSaveFile.Click += new System.EventHandler(this.btnChooseSaveFile_Click);
            // 
            // txtFileName
            // 
            this.txtFileName.Location = new System.Drawing.Point(135, 68);
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.Size = new System.Drawing.Size(397, 20);
            this.txtFileName.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(84, 73);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(45, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Tên File";
            // 
            // btnCreateFileExcel
            // 
            this.btnCreateFileExcel.Location = new System.Drawing.Point(202, 105);
            this.btnCreateFileExcel.Name = "btnCreateFileExcel";
            this.btnCreateFileExcel.Size = new System.Drawing.Size(139, 23);
            this.btnCreateFileExcel.TabIndex = 5;
            this.btnCreateFileExcel.Text = "Tạo File";
            this.btnCreateFileExcel.UseVisualStyleBackColor = true;
            this.btnCreateFileExcel.Click += new System.EventHandler(this.btnCreateFileExcel_Click);
            // 
            // btnChooseFolder
            // 
            this.btnChooseFolder.Location = new System.Drawing.Point(182, 177);
            this.btnChooseFolder.Name = "btnChooseFolder";
            this.btnChooseFolder.Size = new System.Drawing.Size(143, 23);
            this.btnChooseFolder.TabIndex = 1;
            this.btnChooseFolder.Text = "Chọn thư mục";
            this.btnChooseFolder.UseVisualStyleBackColor = true;
            this.btnChooseFolder.Click += new System.EventHandler(this.btnChooseFolder_Click);
            // 
            // btnReChooseFile
            // 
            this.btnReChooseFile.Location = new System.Drawing.Point(386, 105);
            this.btnReChooseFile.Name = "btnReChooseFile";
            this.btnReChooseFile.Size = new System.Drawing.Size(75, 23);
            this.btnReChooseFile.TabIndex = 6;
            this.btnReChooseFile.Text = "Chọn lại file";
            this.btnReChooseFile.UseVisualStyleBackColor = true;
            this.btnReChooseFile.Click += new System.EventHandler(this.btnReChooseFile_Click);
            // 
            // btnCopyFilter
            // 
            this.btnCopyFilter.Location = new System.Drawing.Point(73, 232);
            this.btnCopyFilter.Name = "btnCopyFilter";
            this.btnCopyFilter.Size = new System.Drawing.Size(75, 23);
            this.btnCopyFilter.TabIndex = 7;
            this.btnCopyFilter.Text = "Copy Filter";
            this.btnCopyFilter.UseVisualStyleBackColor = true;
            this.btnCopyFilter.Click += new System.EventHandler(this.btnCopyFilter_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 288);
            this.Controls.Add(this.btnCopyFilter);
            this.Controls.Add(this.btnReChooseFile);
            this.Controls.Add(this.btnCreateFileExcel);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtFileName);
            this.Controls.Add(this.btnChooseSaveFile);
            this.Controls.Add(this.btnChooseFolder);
            this.Controls.Add(this.btnChoiceMutipleFile);
            this.Name = "Form1";
            this.Text = "Tổng Hợp File Excel";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnChoiceMutipleFile;
        private System.Windows.Forms.Button btnChooseSaveFile;
        private System.Windows.Forms.TextBox txtFileName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnCreateFileExcel;
        private System.Windows.Forms.Button btnChooseFolder;
        private System.Windows.Forms.Button btnReChooseFile;
        private System.Windows.Forms.Button btnCopyFilter;
    }
}

