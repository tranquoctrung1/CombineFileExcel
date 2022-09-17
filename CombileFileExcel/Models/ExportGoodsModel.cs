using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace CombileFileExcel.Models
{
    public class ExportGoodsModel
    {
        public string CustomerId { get; set; }  
        public string CustomerName { get; set; }
        public string Date { get; set; }
        public string GoodsId { get; set; }
        public string GoodsName { get; set; }
        public string Unit { get; set; }
        public string Amount { get; set; }
        public string Price { get; set; }
        public string TotalPrice { get; set; }
        public string Note { get; set; }
        public string OrderNumber { get; set; }
    }
}
