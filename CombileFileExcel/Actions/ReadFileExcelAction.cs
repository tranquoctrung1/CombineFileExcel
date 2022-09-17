using CombileFileExcel.Models;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace CombileFileExcel.Actions
{
    public class ReadFileExcelAction
    {
        public List<CustomerModel> ReadCustomerSheet(string path)
        {
            List<CustomerModel> list = new List<CustomerModel>();
            ValidateRowAction validateRowAction = new ValidateRowAction();

            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];

                Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

                int rows = xlWorksheet.Rows.Count;

                for (int i = 2; i <= rows; i++)
                {
                    if (validateRowAction.ValidateRowCustomer(i, xlRange) == true)
                    {
                        CustomerModel el = new CustomerModel();

                        el.CustomerId = xlRange.Cells[i, 1].Value2 == null ? "" : xlRange.Cells[i, 1].Value2.ToString();
                        el.CustomerName = xlRange.Cells[i, 2].Value2 == null ? "" : xlRange.Cells[i, 2].Value2.ToString();
                        el.StafffId = xlRange.Cells[i, 3].Value2 == null ? "" : xlRange.Cells[i, 3].Value2.ToString();
                        el.StaffName = xlRange.Cells[i, 4].Value2 == null ? "" : xlRange.Cells[i, 4].Value2.ToString();

                        list.Add(el);
                    }
                    else
                    {
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return list;
        }

        public void WriteSheetCustomer(List<CustomerModel> list, string path)
        {
            if(path == "")
            {
                MessageBox.Show("Chưa tạo file!!");
            }
            else
            {
                Application ExcelApp = new Application();

                Workbook ExcelWorkBook = null;

                Worksheet ExcelWorkSheet = null;


                try
                {
                    ExcelWorkBook = ExcelApp.Workbooks.Open(path);
                    ExcelWorkSheet = ExcelWorkBook.Worksheets[1];


                    ExcelWorkSheet.Cells[1, 1].Value2 = "MÃ KHÁCH HÀNG";
                    ExcelWorkSheet.Cells[1, 2].Value2 = "TÊN KHÁCH HÀNG";
                    ExcelWorkSheet.Cells[1, 3].Value2 = "MÃ NHÂN VIÊN";
                    ExcelWorkSheet.Cells[1, 4].Value2 = "TÊN NHÂN VIÊN";

                    int indexRow = 2;

                    foreach (CustomerModel customer in list)
                    {
                        ExcelWorkSheet.Cells[indexRow, 1] = customer.CustomerId;
                        ExcelWorkSheet.Cells[indexRow, 2] = customer.CustomerName;
                        ExcelWorkSheet.Cells[indexRow, 3] = customer.StafffId;
                        ExcelWorkSheet.Cells[indexRow, 4] = customer.StaffName;

                        indexRow += 1;
                    }
                    ExcelWorkBook.Save();

                    MessageBox.Show("Write Sheet 1 Done");

                    ExcelWorkBook.Close();

                    ExcelApp.Quit();

                    Marshal.ReleaseComObject(ExcelWorkSheet);

                    Marshal.ReleaseComObject(ExcelWorkBook);

                    Marshal.ReleaseComObject(ExcelApp);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {

                }
            }
        }

        public void CreateFileExcel(string path)
        {
            Application ExcelApp = new Application();

            Workbook ExcelWorkBook = null;

            Worksheet ExcelWorkSheet = null;

            ExcelApp.Visible = false;

            ExcelWorkBook = ExcelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

            try
            {
                var xlSheets = ExcelWorkBook.Sheets as Microsoft.Office.Interop.Excel.Sheets;
                var xlNewSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);
                xlNewSheet.Name = "MÃ KH";

                var xlNewSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)xlSheets.Add(xlSheets[2], Type.Missing, Type.Missing, Type.Missing);
                xlNewSheet2.Name = "MÃ HH";

                var xlNewSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)xlSheets.Add(xlSheets[3], Type.Missing, Type.Missing, Type.Missing);
                xlNewSheet3.Name = "TỔNG HỢP";

                var xlNewSheet4 = (Microsoft.Office.Interop.Excel.Worksheet)xlSheets.Add(xlSheets[4], Type.Missing, Type.Missing, Type.Missing);
                xlNewSheet4.Name = "NHẬP KHẨU";

                var xlNewSheet5 = (Microsoft.Office.Interop.Excel.Worksheet)xlSheets.Add(xlSheets[5], Type.Missing, Type.Missing, Type.Missing);
                xlNewSheet5.Name = "XUẤT KHẨU";

                var xlNewSheet6 = (Microsoft.Office.Interop.Excel.Worksheet)xlSheets.Add(xlSheets[6], Type.Missing, Type.Missing, Type.Missing);
                xlNewSheet6.Name = "Sheet 5";

                ExcelWorkSheet = ExcelWorkBook.Worksheets[1];

                ExcelWorkSheet.Cells[1, 1] = "create file";

                ExcelWorkBook.SaveAs(path);

                ExcelWorkBook.Close();

                ExcelApp.Quit();

                Marshal.ReleaseComObject(ExcelWorkSheet);

                Marshal.ReleaseComObject(ExcelWorkBook);

                Marshal.ReleaseComObject(ExcelApp);

                MessageBox.Show("Tạo file excel thành công!!!");

            }

            catch (Exception exHandle)

            {

                Console.WriteLine("Exception: " + exHandle.Message);

                Console.ReadLine();

            }

            finally
            {
            }
        }
    }
}
