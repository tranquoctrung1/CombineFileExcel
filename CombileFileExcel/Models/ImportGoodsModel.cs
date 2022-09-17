using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CombileFileExcel.Models
{
    public class ImportGoodsModel
    {
        public string CustomerId { get; set; }
        public string CustomerName { get; set; }
        public string  Date { get; set; }
        public string GoodsId { get; set; }
        public string GoodsName { get; set; }
        public string ImportTL { get; set; }
        public string Price { get; set; }
        public string TotalPrice { get; set; }
        public string TL { get; set; }
        public string Note { get; set; }

    }
}
