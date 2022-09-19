using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.Eventing.Reader;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace CombileFileExcel.Actions
{
    public class ValidateRowAction
    {
        public bool ValidateRowCustomer(int rowIndex, Microsoft.Office.Interop.Excel.Range xlRange)
        {
            bool check = true;

            for (int i = 1; i <= 4; i++)
            {
                if (xlRange.Cells[rowIndex, i] == null || xlRange.Cells[rowIndex, i].Value2 == null)
                {
                    check = false;
                }
                else
                {
                    check = true;
                    break;
                }
            }

            return check;
        }

        public bool ValidateRowGoods(int rowIndex, Microsoft.Office.Interop.Excel.Range xlRange)
        {
            bool check = true;

            for (int i = 1; i <= 3; i++)
            {
                if (xlRange.Cells[rowIndex, i] == null || xlRange.Cells[rowIndex, i].Value2 == null)
                {
                    check = false;
                }
                else
                {
                    check = true;
                    break;
                }
            }

            return check;
        }

        public bool ValidateRowTotal(int rowIndex, Microsoft.Office.Interop.Excel.Range xlRange)
        {
            bool check = true;

            for (int i = 1; i <= 18; i++)
            {
                if (xlRange.Cells[rowIndex, i] == null || xlRange.Cells[rowIndex, i].Value2 == null)
                {
                    check = false;
                }
                else
                {
                    check = true;
                    break;
                }
            }

            return check;
        }

        public bool ValidateRowImportGoods(int rowIndex, Microsoft.Office.Interop.Excel.Range xlRange)
        {
            bool check = true;

            for (int i = 1; i <= 10; i++)
            {
                if (xlRange.Cells[rowIndex, i] == null || xlRange.Cells[rowIndex, i].Value2 == null)
                {
                    check = false;
                }
                else
                {
                    check = true;
                    break;
                }
            }

            return check;
        }

        public bool ValidateRowExportGoods(int rowIndex, Microsoft.Office.Interop.Excel.Range xlRange)
        {
            bool check = true;

            for (int i = 1; i <= 11; i++)
            {
                if (xlRange.Cells[rowIndex, i] == null || xlRange.Cells[rowIndex, i].Value2 == null)
                {
                    check = false;
                }
                else
                {
                    check = true;
                    break;
                }
            }

            return check;
        }
    }
}
