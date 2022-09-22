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
            this.btnFilterImportGoods = new System.Windows.Forms.Button();
            this.fileSystemWatcher1 = new System.IO.FileSystemWatcher();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label3 = new System.Windows.Forms.Label();
            this.btnSumSheetImportGoods = new System.Windows.Forms.Button();
            this.btnSumSheetExportGoods = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.txtPathFileCreate = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.txtPathFileExecute = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnChoiceMutipleFile
            // 
            this.btnChoiceMutipleFile.Location = new System.Drawing.Point(193, 151);
            this.btnChoiceMutipleFile.Name = "btnChoiceMutipleFile";
            this.btnChoiceMutipleFile.Size = new System.Drawing.Size(122, 23);
            this.btnChoiceMutipleFile.TabIndex = 0;
            this.btnChoiceMutipleFile.Text = "Chọn file lấy dữ liệu";
            this.btnChoiceMutipleFile.UseVisualStyleBackColor = true;
            this.btnChoiceMutipleFile.Click += new System.EventHandler(this.btnChoiceMutipleFile_Click);
            // 
            // btnChooseSaveFile
            // 
            this.btnChooseSaveFile.Location = new System.Drawing.Point(444, 110);
            this.btnChooseSaveFile.Name = "btnChooseSaveFile";
            this.btnChooseSaveFile.Size = new System.Drawing.Size(117, 23);
            this.btnChooseSaveFile.TabIndex = 2;
            this.btnChooseSaveFile.Text = "Chọn nơi lưu file";
            this.btnChooseSaveFile.UseVisualStyleBackColor = true;
            this.btnChooseSaveFile.Click += new System.EventHandler(this.btnChooseSaveFile_Click);
            // 
            // txtFileName
            // 
            this.txtFileName.Location = new System.Drawing.Point(97, 110);
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.Size = new System.Drawing.Size(330, 20);
            this.txtFileName.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(46, 113);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(45, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Tên File";
            // 
            // btnCreateFileExcel
            // 
            this.btnCreateFileExcel.Location = new System.Drawing.Point(580, 110);
            this.btnCreateFileExcel.Name = "btnCreateFileExcel";
            this.btnCreateFileExcel.Size = new System.Drawing.Size(117, 23);
            this.btnCreateFileExcel.TabIndex = 5;
            this.btnCreateFileExcel.Text = "Tạo File gộp";
            this.btnCreateFileExcel.UseVisualStyleBackColor = true;
            this.btnCreateFileExcel.Click += new System.EventHandler(this.btnCreateFileExcel_Click);
            // 
            // btnChooseFolder
            // 
            this.btnChooseFolder.Location = new System.Drawing.Point(395, 151);
            this.btnChooseFolder.Name = "btnChooseFolder";
            this.btnChooseFolder.Size = new System.Drawing.Size(143, 23);
            this.btnChooseFolder.TabIndex = 1;
            this.btnChooseFolder.Text = "Chọn thư mục lấy dữ liệu";
            this.btnChooseFolder.UseVisualStyleBackColor = true;
            this.btnChooseFolder.Click += new System.EventHandler(this.btnChooseFolder_Click);
            // 
            // btnReChooseFile
            // 
            this.btnReChooseFile.Location = new System.Drawing.Point(352, 100);
            this.btnReChooseFile.Name = "btnReChooseFile";
            this.btnReChooseFile.Size = new System.Drawing.Size(75, 23);
            this.btnReChooseFile.TabIndex = 6;
            this.btnReChooseFile.Text = "Chọn lại file";
            this.btnReChooseFile.UseVisualStyleBackColor = true;
            this.btnReChooseFile.Click += new System.EventHandler(this.btnReChooseFile_Click);
            // 
            // btnCopyFilter
            // 
            this.btnCopyFilter.Location = new System.Drawing.Point(100, 151);
            this.btnCopyFilter.Name = "btnCopyFilter";
            this.btnCopyFilter.Size = new System.Drawing.Size(268, 23);
            this.btnCopyFilter.TabIndex = 7;
            this.btnCopyFilter.Text = "Chuyển dữ liệu trang nhập hàng sang trang xuất hàng";
            this.btnCopyFilter.UseVisualStyleBackColor = true;
            this.btnCopyFilter.Click += new System.EventHandler(this.btnCopyFilter_Click);
            // 
            // btnFilterImportGoods
            // 
            this.btnFilterImportGoods.Location = new System.Drawing.Point(413, 151);
            this.btnFilterImportGoods.Name = "btnFilterImportGoods";
            this.btnFilterImportGoods.Size = new System.Drawing.Size(273, 23);
            this.btnFilterImportGoods.TabIndex = 8;
            this.btnFilterImportGoods.Text = "Lọc dữ liệu trang nhập hàng";
            this.btnFilterImportGoods.UseVisualStyleBackColor = true;
            this.btnFilterImportGoods.Click += new System.EventHandler(this.btnFilterImportGoods_Click);
            // 
            // fileSystemWatcher1
            // 
            this.fileSystemWatcher1.EnableRaisingEvents = true;
            this.fileSystemWatcher1.SynchronizingObject = this;
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.txtPathFileCreate);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.btnChooseSaveFile);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.txtFileName);
            this.panel1.Controls.Add(this.btnCreateFileExcel);
            this.panel1.Controls.Add(this.btnChoiceMutipleFile);
            this.panel1.Controls.Add(this.btnChooseFolder);
            this.panel1.Location = new System.Drawing.Point(12, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(776, 209);
            this.panel1.TabIndex = 9;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(229, 15);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(319, 31);
            this.label2.TabIndex = 5;
            this.label2.Text = "Gộp nhiều file thành 1 file";
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel2.Controls.Add(this.txtPathFileExecute);
            this.panel2.Controls.Add(this.label5);
            this.panel2.Controls.Add(this.btnSumSheetExportGoods);
            this.panel2.Controls.Add(this.btnSumSheetImportGoods);
            this.panel2.Controls.Add(this.label3);
            this.panel2.Controls.Add(this.btnCopyFilter);
            this.panel2.Controls.Add(this.btnFilterImportGoods);
            this.panel2.Controls.Add(this.btnReChooseFile);
            this.panel2.Location = new System.Drawing.Point(12, 239);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(775, 247);
            this.panel2.TabIndex = 10;
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(264, 17);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(284, 31);
            this.label3.TabIndex = 6;
            this.label3.Text = "Thao tác tên các trang";
            // 
            // btnSumSheetImportGoods
            // 
            this.btnSumSheetImportGoods.Location = new System.Drawing.Point(100, 193);
            this.btnSumSheetImportGoods.Name = "btnSumSheetImportGoods";
            this.btnSumSheetImportGoods.Size = new System.Drawing.Size(268, 23);
            this.btnSumSheetImportGoods.TabIndex = 9;
            this.btnSumSheetImportGoods.Text = "Tính tổng trang nhập hàng";
            this.btnSumSheetImportGoods.UseVisualStyleBackColor = true;
            this.btnSumSheetImportGoods.Click += new System.EventHandler(this.btnSumSheetImportGoods_Click);
            // 
            // btnSumSheetExportGoods
            // 
            this.btnSumSheetExportGoods.Location = new System.Drawing.Point(413, 193);
            this.btnSumSheetExportGoods.Name = "btnSumSheetExportGoods";
            this.btnSumSheetExportGoods.Size = new System.Drawing.Size(273, 23);
            this.btnSumSheetExportGoods.TabIndex = 10;
            this.btnSumSheetExportGoods.Text = "Tính tông trang xuất hàng";
            this.btnSumSheetExportGoods.UseVisualStyleBackColor = true;
            this.btnSumSheetExportGoods.Click += new System.EventHandler(this.btnSumSheetExportGoods_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(49, 60);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(97, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Đường dẫn file tạo:";
            // 
            // txtPathFileCreate
            // 
            this.txtPathFileCreate.AutoSize = true;
            this.txtPathFileCreate.Location = new System.Drawing.Point(152, 60);
            this.txtPathFileCreate.Name = "txtPathFileCreate";
            this.txtPathFileCreate.Size = new System.Drawing.Size(0, 13);
            this.txtPathFileCreate.TabIndex = 7;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(100, 69);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(146, 13);
            this.label5.TabIndex = 11;
            this.label5.Text = "Đường dẫn file đang tao tác: ";
            // 
            // txtPathFileExecute
            // 
            this.txtPathFileExecute.AutoSize = true;
            this.txtPathFileExecute.Location = new System.Drawing.Point(252, 69);
            this.txtPathFileExecute.Name = "txtPathFileExecute";
            this.txtPathFileExecute.Size = new System.Drawing.Size(0, 13);
            this.txtPathFileExecute.TabIndex = 12;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 517);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "Form1";
            this.Text = "Tổng Hợp File Excel";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);

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
        private System.Windows.Forms.Button btnFilterImportGoods;
        private System.IO.FileSystemWatcher fileSystemWatcher1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnSumSheetImportGoods;
        private System.Windows.Forms.Button btnSumSheetExportGoods;
        private System.Windows.Forms.Label txtPathFileExecute;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label txtPathFileCreate;
        private System.Windows.Forms.Label label4;
    }
}

