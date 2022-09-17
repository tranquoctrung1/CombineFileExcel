using CombileFileExcel.Actions;
using CombileFileExcel.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CombileFileExcel
{
    public partial class Form1 : Form
    {
        string pathToSaveFile;
        string fileName;
        string path;
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
            List<CustomerModel> listCustomer = new List<CustomerModel>();

            ReadFileExcelAction readFileExcelAction = new ReadFileExcelAction();

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel (*.xls; *.xlsx) | *.xls; *.xlsx|" + "All files (*.*)|*.*";

            openFileDialog.Multiselect = true;
            openFileDialog.Title = "Chọn nhiều file excel";

            DialogResult dr = openFileDialog.ShowDialog();

            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                // Read the files
                foreach (String file in openFileDialog.FileNames)
                {
                    try
                    {
                        List<CustomerModel> temp = readFileExcelAction.ReadCustomerSheet(file);
                        if(temp.Count > 0)
                        {
                            listCustomer.AddRange(temp);
                        }
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

                if(listCustomer.Count > 0)
                {
                    readFileExcelAction.WriteSheetCustomer(listCustomer, path);
                }
            }
        }

        private void btnChooseFolder_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    List<string> files = Directory.GetFiles(fbd.SelectedPath).Where(s => s.EndsWith(".xls") || s.EndsWith(".xlsx")).ToList();

                    foreach(string file in files)
                    {
                        MessageBox.Show(file);
                    }
                }
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
            if(txtFileName.Text == "")
            {
                MessageBox.Show("Phải đặt tên file excel!!!");
                return;
            }
            else if(pathToSaveFile == "")
            {
                MessageBox.Show("Phải chọn nơi lưu file!!!");
                return;
            }
            else
            {
                fileName = txtFileName.Text;
                path = $"{pathToSaveFile}{fileName}.xlsx";

                if(!File.Exists(path))
                {
                    ReadFileExcelAction readFileExcelAction = new ReadFileExcelAction();

                    readFileExcelAction.CreateFileExcel(path);
                }
                else
                {
                    MessageBox.Show("File đã tồn tại!!!. Hãy tạo mới 1 file ");
                }
            }

        }
    }
}
