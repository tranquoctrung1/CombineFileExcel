using CombileFileExcel.Models;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO; 
using System.Linq;
using System.Net.Http;
using System.Reflection;
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
                xlNewSheet4.Name = "NHẬP HÀNG";

                var xlNewSheet5 = (Microsoft.Office.Interop.Excel.Worksheet)xlSheets.Add(xlSheets[5], Type.Missing, Type.Missing, Type.Missing);
                xlNewSheet5.Name = "XUẤT HÀNG";

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

        public List<ImportGoodsModel> LoadFileExcelSheetImportGoods(string path)
        {

            List<ImportGoodsModel> list = new List<ImportGoodsModel>();

            try
            {
                string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;";
                System.Data.OleDb.OleDbConnection MyConnection;
                System.Data.DataSet DtSet;
                System.Data.OleDb.OleDbDataAdapter MyCommand;
                MyConnection = new System.Data.OleDb.OleDbConnection(connStr);
                MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [NHẬP HÀNG$]", MyConnection);
                MyCommand.TableMappings.Add("Table", "Net-informations.com");
                DtSet = new System.Data.DataSet();
                MyCommand.Fill(DtSet);
                MyConnection.Close();

                if (DtSet.Tables[0].Rows.Count > 0)
                {
                    foreach(DataRow row in DtSet.Tables[0].Rows)
                    {
                        if (row[0] != null)
                        {
                            if (row[0].ToString() != "")
                            {
                                if (row[0].ToString().ToLower() != "kho3" && row[0].ToString().ToLower()  != "kho4" && row[0].ToString().ToLower() != "vinhkhanh" && row[0].ToString().ToLower() != "kho ncq" && row[0].ToString().ToLower() != "tổng")
                                {
                                    ImportGoodsModel el = new ImportGoodsModel();

                                    el.CustomerID = row[0].ToString();
                                    el.CustomerName = row[1].ToString();

                                    el.TimeStamp = row[2].ToString();
                                    el.GoodsName = row[4].ToString();

                                    if (row[13].ToString() == "")
                                    {
                                        el.Amout = row[13].ToString(); 
                                    }
                                    else
                                    {
                                        el.Amout = "-" + row[13].ToString();
                                    }
                                    el.Price = row[14].ToString();
                                    el.TotalPrice = row[15].ToString();

                                    list.Add(el);
                                }

                            }
                        }
                    }

                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return list;
        }

        public System.Data.DataTable LoadFileExcelSheetImportGoodToDataTabe(string path)
        {
            System.Data.DataTable table = new System.Data.DataTable();
            try
            {
                string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;";
                System.Data.OleDb.OleDbConnection MyConnection;
                System.Data.DataSet DtSet;
                System.Data.OleDb.OleDbDataAdapter MyCommand;
                MyConnection = new System.Data.OleDb.OleDbConnection(connStr);
                MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [NHẬP HÀNG$]", MyConnection);
                MyCommand.TableMappings.Add("Table", "Net-informations.com");
                DtSet = new System.Data.DataSet();
                MyCommand.Fill(DtSet);
                MyConnection.Close();

                if (DtSet.Tables.Count > 0)
                {

                    table = DtSet.Tables[0];
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return table;
        }


        public int GetUsedRowExcelInExportGoodsSheet(string path)
        {
            int length = 0;

            try
            {
                string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;";
                System.Data.OleDb.OleDbConnection MyConnection;
                System.Data.DataSet DtSet;
                System.Data.OleDb.OleDbDataAdapter MyCommand;
                MyConnection = new System.Data.OleDb.OleDbConnection(connStr);
                MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [XUẤT HÀNG$]", MyConnection);
                MyCommand.TableMappings.Add("Table", "Net-informations.com");
                DtSet = new System.Data.DataSet();
                MyCommand.Fill(DtSet);
                MyConnection.Close();

                if (DtSet.Tables[0].Rows.Count > 0)
                {
                    length = DtSet.Tables[0].Rows.Count;
                }
            }
            catch(Exception ex)
            {

            }
            finally
            {

            }

            return length;
        }

        public int GetUsedRowExcelInImportGoodsSheet(string path)
        {
            int length = 0;

            try
            {
                string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;";
                System.Data.OleDb.OleDbConnection MyConnection;
                System.Data.DataSet DtSet;
                System.Data.OleDb.OleDbDataAdapter MyCommand;
                MyConnection = new System.Data.OleDb.OleDbConnection(connStr);
                MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [NHẬP HÀNG$]", MyConnection);
                MyCommand.TableMappings.Add("Table", "Net-informations.com");
                DtSet = new System.Data.DataSet();
                MyCommand.Fill(DtSet);
                MyConnection.Close();

                if (DtSet.Tables[0].Rows.Count > 0)
                {
                    length = DtSet.Tables[0].Rows.Count;
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {

            }

            return length;
        }


        public void WriteToExportGoods(List<ImportGoodsModel> list, string path)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[5];

                int usedRow = GetUsedRowExcelInExportGoodsSheet(path);
                if(usedRow != 0)
                {
                    usedRow += 2;
                }

                if(list.Count > 0)
                {
                    foreach(ImportGoodsModel item in list)
                    {
                        xlWorksheet.Cells[usedRow, 1] = item.CustomerID ?? "";
                        xlWorksheet.Cells[usedRow, 2] = item.CustomerName ?? "";
                        xlWorksheet.Cells[usedRow, 3] = item.TimeStamp ?? "";
                        xlWorksheet.Cells[usedRow, 5] = item.GoodsName ?? "";
                        xlWorksheet.Cells[usedRow, 7] = item.Amout ?? "";
                        xlWorksheet.Cells[usedRow, 8] = item.Price ?? "";
                        xlWorksheet.Cells[usedRow, 9] = item.TotalPrice ?? "";

                        usedRow++;
                    }
                }

                xlWorkbook.Close(true);

                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorksheet);
                Marshal.ReleaseComObject(xlWorkbook);

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
        }



        public void WriteToExportGoodsByDataTable(string path, System.Data.DataTable tableData)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[4];

                xlWorksheet.Rows.Clear();
                xlWorksheet.Columns.Clear();

                for (var i = 0; i < tableData.Columns.Count; i++)
                {
                    xlWorksheet.Cells[1, i + 1] = tableData.Columns[i].ColumnName;
                }

                int rowIndex = 2;
                int columnIndex = 1;

                for (int  i  = 0; i < tableData.Rows.Count; i++)
                {
                    if (tableData.Rows[i][0].ToString() != "")
                    {
                        if (tableData.Rows[i][0].ToString().ToLower() == "kho3" || tableData.Rows[i][0].ToString().ToLower() == "kho4" || tableData.Rows[i][0].ToString().ToLower() == "vinhkhanh" || tableData.Rows[i][0].ToString().ToLower() == "kho ncq")
                        {
                            columnIndex = 1;
                            for (int j = 0; j < tableData.Columns.Count; j++)
                            {
                                xlWorksheet.Cells[rowIndex, columnIndex++] = tableData.Rows[i][j] ?? "";
                            }
                            rowIndex += 1;
                        }
                    }
                }

                xlWorkbook.Close(true);

                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorksheet);
                Marshal.ReleaseComObject(xlWorkbook);

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
        }

        public void SumSheetImportGoods(string path)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[4];

                int usedRow = GetUsedRowExcelInImportGoodsSheet(path);
                if(usedRow > 0)
                {
                    usedRow += 2;
                }

                if (xlWorksheet.Cells[usedRow - 1, 1] != null)
                {
                    if(xlWorksheet.Cells[usedRow - 1, 1].Value == "TỔNG")
                    {
                        Range r = xlWorksheet.Range[xlWorksheet.Cells[usedRow - 1, 1], xlWorksheet.Cells[usedRow - 1, 18]];
                        r.EntireRow.Delete(XlDeleteShiftDirection.xlShiftUp);

                        usedRow -= 1;
                    }
                }

                xlWorksheet.Cells[usedRow, 1] = "TỔNG";
                xlWorksheet.Range[xlWorksheet.Cells[usedRow, 1], xlWorksheet.Cells[usedRow, 7]].Merge();
                xlWorksheet.Cells[usedRow, 8].Formula = string.Format("=SUBTOTAL(9,H2:H{0})", usedRow - 1);
                xlWorksheet.Cells[usedRow, 9].Formula = string.Format("=SUBTOTAL(9,I2:I{0})", usedRow - 1);
                xlWorksheet.Cells[usedRow, 14].Formula = string.Format("=SUBTOTAL(9,N2:N{0})", usedRow - 1);
                xlWorksheet.Cells[usedRow, 15].Formula = string.Format("=SUBTOTAL(9,O2:O{0})", usedRow - 1);
                xlWorksheet.Cells[usedRow, 16].Formula = string.Format("=SUBTOTAL(9,P2:P{0})", usedRow - 1);

                xlWorkbook.Close(true);

                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorksheet);
                Marshal.ReleaseComObject(xlWorkbook);

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
        }

        public void SumSheetExportGoods(string path)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[5];

                int usedRow = GetUsedRowExcelInExportGoodsSheet(path);
                if (usedRow > 0)
                {
                    usedRow += 2;
                }

                if (xlWorksheet.Cells[usedRow - 1, 1] != null)
                {
                    if (xlWorksheet.Cells[usedRow - 1, 1].Value == "TỔNG")
                    {
                        Range r = xlWorksheet.Range[xlWorksheet.Cells[usedRow - 1, 1], xlWorksheet.Cells[usedRow - 1, 18]];
                        r.EntireRow.Delete(XlDeleteShiftDirection.xlShiftUp);

                        usedRow -= 1;
                    }
                }

                xlWorksheet.Cells[usedRow, 1] = "TỔNG";
                xlWorksheet.Range[xlWorksheet.Cells[usedRow, 1], xlWorksheet.Cells[usedRow, 6]].Merge();
                xlWorksheet.Cells[usedRow, 7].Formula = string.Format("=SUBTOTAL(9,G2:G{0})", usedRow - 1);
                xlWorksheet.Cells[usedRow, 8].Formula = string.Format("=SUBTOTAL(9,H2:H{0})", usedRow - 1);
                xlWorksheet.Cells[usedRow, 10].Formula = string.Format("=SUBTOTAL(9,J2:J{0})", usedRow - 1);
                xlWorksheet.Cells[usedRow, 14].Formula = string.Format("=SUBTOTAL(9,N2:N{0})", usedRow - 1);
                xlWorksheet.Cells[usedRow, 15].Formula = string.Format("=SUBTOTAL(9,O2:O{0})", usedRow - 1);
                xlWorksheet.Cells[usedRow, 16].Formula = string.Format("=SUBTOTAL(9,P2:P{0})", usedRow - 1);

                xlWorkbook.Close(true);

                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorksheet);
                Marshal.ReleaseComObject(xlWorkbook);

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
        }

    }
}
