using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CombileFileExcel.Models
{
    public class ImportGoodsModel
    {
       public string CustomerID { get; set; }
        public string CustomerName { get; set; }
        public string TimeStamp { get; set; }
        public string GoodsName { get; set; }
        public string Amout { get; set; }
        public string Price { get; set; }
        public string TotalPrice { get; set; }
        

    }
}
