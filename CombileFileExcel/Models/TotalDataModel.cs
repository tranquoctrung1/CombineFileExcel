using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CombileFileExcel.Models
{
    public class TotalDataModel
    {
        public List<CustomerModel> ListCustomer { get; set; }
        public List<GoodsModel> ListGoods { get; set; }
        public List<TotalModel> ListTotal { get; set; }
        public List<ImportGoodsModel> ListImportGoods { get; set; }
        public List<ExportGoodsModel> ListExportGoods { get; set; }
    }
}
