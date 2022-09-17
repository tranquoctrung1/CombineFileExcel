using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace CombileFileExcel.Models
{
    public  class TotalModel
    {
        public string GoodsId { get; set; }
        public string GoodsName { get; set; }
        public string RemainStartMonth { get; set; }
        public string ImportVK { get; set; }
        public string ImportNCQ { get; set; }
        public string ImportFMVN { get; set; }
        public string ImportSW { get; set; }
        public string ImportCLK { get; set; }
        public string ImportTL { get; set; }
        public string ChangeShell { get; set; }
        public string ExportToSell { get; set; }
        public string ExportToTranmission { get; set; }

        public string RemainEndMonth { get; set; }
        public string MiniStock { get; set; }
        public string Deviant { get; set; }
        public string Note { get; set; }
        public string Status { get; set; }

    }
}
