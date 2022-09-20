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
        public int CopyCustomerSheet(string path, string pathToFileSave, int index)
        {
            int result = 0;
            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];

                Microsoft.Office.Interop.Excel.Workbook xlWorkbook2 = xlApp.Workbooks.Open(pathToFileSave, 0, false);
                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet2 = xlWorkbook2.Sheets[1];

                if (xlWorksheet.AutoFilter != null)
                {
                    xlWorksheet.AutoFilterMode = false;
                }

                xlWorksheet.Rows.ClearFormats();
                xlWorksheet.Columns.ClearFormats();

                result = xlWorksheet.UsedRange.Rows.Count;

                int indexGetContent = 2;
                int indexFillConent = index + 1;

                if(index ==  1)
                {
                    indexGetContent = 1;
                    indexFillConent = 1;
                }

                Microsoft.Office.Interop.Excel.Range from = xlWorksheet.Range[$"A{indexGetContent}:D{result}"];
                Microsoft.Office.Interop.Excel.Range to = xlWorksheet2.Range[$"A{indexFillConent}:D{result + index}"];

                from.Copy(to);

                xlWorkbook2.Save();
                xlWorksheet2.Columns.AutoFit();

                xlWorkbook.Close(false);

                xlWorkbook2.Close(true);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorksheet);
                Marshal.ReleaseComObject(xlWorkbook);

                Marshal.ReleaseComObject(xlWorksheet2);
                Marshal.ReleaseComObject(xlWorkbook2);

                Marshal.ReleaseComObject(xlApp);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                foreach (Process process in Process.GetProcessesByName("Excel"))
                {
                    process.Kill();
                }
            }

            return result;
        }

        public int CopyGoodsSheet(string path, string pathToFileSave, int index)
        {
            int result = 0;
            try
            {

                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];

                Microsoft.Office.Interop.Excel.Workbook xlWorkbook2 = xlApp.Workbooks.Open(pathToFileSave, 0, false);
                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet2 = xlWorkbook2.Sheets[2];

                if (xlWorksheet.AutoFilter != null)
                {
                    xlWorksheet.AutoFilterMode = false;
                }

                xlWorksheet.Rows.ClearFormats();
                xlWorksheet.Columns.ClearFormats();

                result = xlWorksheet.UsedRange.Rows.Count;

                int indexGetContent = 2;
                int indexFillConent = index + 1;

                if (index == 1)
                {
                    indexGetContent = 1;
                    indexFillConent = 1;
                }

                Microsoft.Office.Interop.Excel.Range from = xlWorksheet.Range[$"A{indexGetContent}:C{result}"];
                Microsoft.Office.Interop.Excel.Range to = xlWorksheet2.Range[$"A{indexFillConent}:C{result + index}"];

                from.Copy(to);

                xlWorkbook2.Save();
                xlWorksheet2.Columns.AutoFit();

                xlWorkbook.Close(false);

                xlWorkbook2.Close(true);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorksheet);
                Marshal.ReleaseComObject(xlWorkbook);

                Marshal.ReleaseComObject(xlWorksheet2);
                Marshal.ReleaseComObject(xlWorkbook2);

                Marshal.ReleaseComObject(xlApp);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                foreach (Process process in Process.GetProcessesByName("Excel"))
                {
                    process.Kill();
                }
            }
            return result;
        }

        public int CopyImportGoods(string path, string pathToFileSave, int index)
        {
            int result = 0;
            try
            {

                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[4];

                Microsoft.Office.Interop.Excel.Workbook xlWorkbook2 = xlApp.Workbooks.Open(pathToFileSave, 0, false);
                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet2 = xlWorkbook2.Sheets[4];

                if (xlWorksheet.AutoFilter != null)
                {
                    xlWorksheet.AutoFilterMode = false;
                }
                xlWorksheet.Rows.ClearFormats();
                xlWorksheet.Columns.ClearFormats();

                result = xlWorksheet.UsedRange.Rows.Count - 1;

                for(int i = result; i >= 0; i--)
                {
                    if (xlWorksheet.UsedRange.Cells[i,1] != null && xlWorksheet.UsedRange.Cells[i, 1].Value != null)
                    {
                        result = i - 1;
                        break;
                    }
                }

                int indexGetContent = 3;
                int indexFillConent = index + 1;

                if (index == 1)
                {
                    indexGetContent = 2;
                    indexFillConent = 1;
                }

                Microsoft.Office.Interop.Excel.Range from = xlWorksheet.Range[$"A{indexGetContent}:R{result}"];
                Microsoft.Office.Interop.Excel.Range to = xlWorksheet2.Range[$"A{indexFillConent}:R{result + index}"];

                from.Copy(to);

                xlWorkbook2.Save();
                xlWorksheet2.Columns.AutoFit();

                xlWorkbook.Close(false);

                xlWorkbook2.Close(true);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorksheet);
                Marshal.ReleaseComObject(xlWorkbook);

                Marshal.ReleaseComObject(xlWorksheet2);
                Marshal.ReleaseComObject(xlWorkbook2);

                Marshal.ReleaseComObject(xlApp);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                foreach (Process process in Process.GetProcessesByName("Excel"))
                {
                    process.Kill();
                }
            }
            return result;
        }

        public int CopyExportGoods(string path, string pathToFileSave, int index)
        {
            int result = 0;
            try
            {

                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[5];

                Microsoft.Office.Interop.Excel.Workbook xlWorkbook2 = xlApp.Workbooks.Open(pathToFileSave, 0, false);
                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet2 = xlWorkbook2.Sheets[5];

                if (xlWorksheet.AutoFilter != null)
                {
                    xlWorksheet.AutoFilterMode = false;
                }
                xlWorksheet.Rows.ClearFormats();
                xlWorksheet.Columns.ClearFormats();

                result = xlWorksheet.UsedRange.Rows.Count - 1;

                for (int i = result; i >= 0; i--)
                {
                    if (xlWorksheet.UsedRange.Cells[i, 1] != null && xlWorksheet.UsedRange.Cells[i, 1].Value != null)
                    {
                        result = i - 1;
                        break;
                    }
                }

                int indexGetContent = 3;
                int indexFillConent = index + 1;

                if (index == 1)
                {
                    indexGetContent = 2;
                    indexFillConent = 1;
                }

                Microsoft.Office.Interop.Excel.Range from = xlWorksheet.Range[$"A{indexGetContent}:R{result}"];
                Microsoft.Office.Interop.Excel.Range to = xlWorksheet2.Range[$"A{indexFillConent}:R{result + index}"];

                from.Copy(to);

                xlWorkbook2.Save();
                xlWorksheet2.Columns.AutoFit();

                xlWorkbook.Close(false);

                xlWorkbook2.Close(true);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorksheet);
                Marshal.ReleaseComObject(xlWorkbook);

                Marshal.ReleaseComObject(xlWorksheet2);
                Marshal.ReleaseComObject(xlWorkbook2);

                Marshal.ReleaseComObject(xlApp);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                foreach (Process process in Process.GetProcessesByName("Excel"))
                {
                    process.Kill();
                }
            }
            return result;
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

                //var xlNewSheet6 = (Microsoft.Office.Interop.Excel.Worksheet)xlSheets.Add(xlSheets[6], Type.Missing, Type.Missing, Type.Missing);
                //xlNewSheet6.Name = "Sheet 5";

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
                foreach (Process process in Process.GetProcessesByName("Excel"))
                {
                    process.Kill();
                }

            }
        }
    }
}
