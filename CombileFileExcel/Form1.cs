using CombileFileExcel.Actions;
using CombileFileExcel.Models;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.Remoting.Messaging;
using System.Security;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CombileFileExcel
{
    public partial class Form1 : Form
    {
        string pathToSaveFile;
        string fileName;
        string path;

        WaitForm waitForm;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.CenterToScreen();
            pathToSaveFile = "";
            fileName = "";
            path = "";
        }

        private void btnChoiceMutipleFile_Click(object sender, EventArgs e)
        {
            if (path != "")
            {
                ReadFileExcelAction readFileExcelAction = new ReadFileExcelAction();

                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel (*.xls; *.xlsx) | *.xls; *.xlsx|" + "All files (*.*)|*.*";

                openFileDialog.Multiselect = true;
                openFileDialog.Title = "Chọn nhiều file excel";

                bool isFirstLoad = true;

                DialogResult dr = openFileDialog.ShowDialog();

                if (dr == System.Windows.Forms.DialogResult.OK)
                {
                    int indexForCustomerSheet = 1;
                    int indexForGoodsSheet = 1;
                    int indexForImportGoods = 1;
                    int indexForExportGoods = 1;

                    using (waitForm = new WaitForm(() =>
                    {
                        foreach (String file in openFileDialog.FileNames)
                        {

                            try
                            {
                                if (isFirstLoad == true)
                                {
                                    indexForCustomerSheet += readFileExcelAction.CopyCustomerSheet(file, path, indexForCustomerSheet) - 1;
                                    indexForGoodsSheet += readFileExcelAction.CopyGoodsSheet(file, path, indexForGoodsSheet) - 1;

                                    isFirstLoad = false;
                                }

                                indexForImportGoods += readFileExcelAction.CopyImportGoods(file, path, indexForImportGoods) - 2;
                                indexForExportGoods += readFileExcelAction.CopyExportGoods(file, path, indexForExportGoods) - 2;
                            }
                            catch (SecurityException ex)
                            {
                                // The user lacks appropriate permissions to read files, discover paths, etc.
                                MessageBox.Show("Security error. Please contact your administrator for details.\n\n" +
                                    "Error message: " + ex.Message + "\n\n" +
                                    "Details (send to Support):\n\n" + ex.StackTrace
                                );
                            }
                            catch (Exception ex)
                            {
                                // Could not load the image - probably related to Windows file system permissions.
                                MessageBox.Show("Cannot display the image: " + file.Substring(file.LastIndexOf('\\'))
                                    + ". You may not have permission to read the file, or " +
                                    "it may be corrupt.\n\nReported error: " + ex.Message);
                            }
                        }

                    }))
                    {
                        waitForm.ShowDialog();
                    }

                   
                }
            }
            else
            {
                MessageBox.Show("Chưa có lưu file đề gộp!!!");
            }

        }

        private void btnChooseFolder_Click(object sender, EventArgs e)
        {
            if (path != "")
            {
                using (var fbd = new FolderBrowserDialog())
                {
                    DialogResult result = fbd.ShowDialog();

                    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                    {
                        List<string> files = Directory.GetFiles(fbd.SelectedPath).Where(s => s.EndsWith(".xls") || s.EndsWith(".xlsx")).ToList();

                        ReadFileExcelAction readFileExcelAction = new ReadFileExcelAction();

                        bool isFirstLoad = true;

                        int indexForCustomerSheet = 1;
                        int indexForGoodsSheet = 1;
                        int indexForImportGoods = 1;
                        int indexForExportGoods = 1;

                        using (waitForm = new WaitForm(() => {
                            foreach (String file in files)
                            {
                                try
                                {
                                    if (isFirstLoad == true)
                                    {
                                        indexForCustomerSheet += readFileExcelAction.CopyCustomerSheet(file, path, indexForCustomerSheet) - 1;
                                        indexForGoodsSheet += readFileExcelAction.CopyGoodsSheet(file, path, indexForGoodsSheet) - 1;

                                        isFirstLoad = false;
                                    }

                                    indexForImportGoods += readFileExcelAction.CopyImportGoods(file, path, indexForImportGoods) - 2;
                                    indexForExportGoods += readFileExcelAction.CopyExportGoods(file, path, indexForExportGoods) - 2;
                                }
                                catch (SecurityException ex)
                                {
                                    // The user lacks appropriate permissions to read files, discover paths, etc.
                                    MessageBox.Show("Security error. Please contact your administrator for details.\n\n" +
                                        "Error message: " + ex.Message + "\n\n" +
                                        "Details (send to Support):\n\n" + ex.StackTrace);
                                }
                                catch (Exception ex)
                                {
                                    // Could not load the image - probably related to Windows file system permissions.
                                    MessageBox.Show("Cannot display the image: " + file.Substring(file.LastIndexOf('\\'))
                                        + ". You may not have permission to read the file, or " +
                                        "it may be corrupt.\n\nReported error: " + ex.Message);
                                }
                            }
                        }))
                        {
                            waitForm.ShowDialog();
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Chưa có lưu file đề gộp!!!");
            }
        }

        private void btnChooseSaveFile_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    pathToSaveFile = fbd.SelectedPath;

                    if (pathToSaveFile[pathToSaveFile.Length - 1] != '\\')
                    {
                        pathToSaveFile += '\\';
                    }
                }
            }
        }

        private void btnCreateFileExcel_Click(object sender, EventArgs e)
        {
            if (txtFileName.Text == "")
            {
                MessageBox.Show("Phải đặt tên file excel!!!");
                return;
            }
            else if (pathToSaveFile == "")
            {
                MessageBox.Show("Phải chọn nơi lưu file!!!");
                return;
            }
            else
            {
                fileName = txtFileName.Text;
                path = $"{pathToSaveFile}{fileName}.xlsx";

                if (!File.Exists(path))
                {
                    ReadFileExcelAction readFileExcelAction = new ReadFileExcelAction();

                    readFileExcelAction.CreateFileExcel(path);

                    txtPathFileCreate.Text = path;
                    txtPathFileExecute.Text = path;
                }
                else
                {
                    MessageBox.Show("File đã tồn tại!!!. Hãy tạo mới 1 file ");

                    txtPathFileCreate.Text = "";
                }
            }

        }

        private void btnReChooseFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel (*.xls; *.xlsx) | *.xls; *.xlsx|" + "All files (*.*)|*.*";
            openFileDialog.Title = "Chọn file excel";

            DialogResult dr = openFileDialog.ShowDialog();

            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                if (openFileDialog.FileName != "")
                {
                    path = openFileDialog.FileName;

                    txtPathFileExecute.Text = path;
                }
            }
        }

        private void btnCopyFilter_Click(object sender, EventArgs e)
        {
            if (path != "")
            {

                using(waitForm = new WaitForm(() => {
                    ReadFileExcelAction readFileExcelAction = new ReadFileExcelAction();

                    List<ImportGoodsModel> list = readFileExcelAction.LoadFileExcelSheetImportGoods(path);

                    readFileExcelAction.WriteToExportGoods(list, path);
                }))
                {
                    waitForm.ShowDialog();
                }
            }
            else
            {
                MessageBox.Show("File chưa có!!!");
            }
        }

        private void btnFilterImportGoods_Click(object sender, EventArgs e)
        {
            if (path != "")
            {

                using(waitForm = new WaitForm(() => {
                    ReadFileExcelAction readFileExcelAction = new ReadFileExcelAction();

                    System.Data.DataTable dt = readFileExcelAction.LoadFileExcelSheetImportGoodToDataTabe(path);

                    readFileExcelAction.WriteToExportGoodsByDataTable(path, dt);
                }))
                {
                    waitForm.ShowDialog();
                }
            }
            else
            {
                MessageBox.Show("File chưa có!!!");
            }

        }

        private void btnSumSheetImportGoods_Click(object sender, EventArgs e)
        {
            if (path != "")
            {

                using(waitForm = new WaitForm(() => {
                    ReadFileExcelAction readFileExcelAction = new ReadFileExcelAction();
                    readFileExcelAction.SumSheetImportGoods(path);

                }))
                {
                    waitForm.ShowDialog();
                }

               
            }
            else
            {
                MessageBox.Show("File chưa có!!!");
            }

        }

        private void btnSumSheetExportGoods_Click(object sender, EventArgs e)
        {
            if (path != "")
            {

                using (waitForm = new WaitForm(() => {
                    ReadFileExcelAction readFileExcelAction = new ReadFileExcelAction();
                    readFileExcelAction.SumSheetExportGoods(path);
                }))
                {
                    waitForm.ShowDialog();
                }
                
            }
            else
            {
                MessageBox.Show("File chưa có!!!");
            }
        }
    }
}
