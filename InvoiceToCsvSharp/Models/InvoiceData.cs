using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InvoiceToCsvSharp.Models
{
    public class InvoiceData
    {
        [CsvHelper.Configuration.Attributes.Name("发票代码")]
        [JsonProperty("发票代码")]
        public string InvoiceCode { get; set; }

        [CsvHelper.Configuration.Attributes.Name("发票号码")]
        [JsonProperty("发票号码")]
        public string InvoiceNumber { get; set; }

        [CsvHelper.Configuration.Attributes.Name("销方识别号")]
        [JsonProperty("销方识别号")]
        public string SellerTaxId { get; set; }

        [CsvHelper.Configuration.Attributes.Name("销方名称")]
        [JsonProperty("销方名称")]
        public string SellerName { get; set; }

        [CsvHelper.Configuration.Attributes.Name("购方识别号")]
        [JsonProperty("购方识别号")]
        public string BuyerTaxId { get; set; }

        [CsvHelper.Configuration.Attributes.Name("购买方名称")]
        [JsonProperty("购买方名称")]
        public string BuyerName { get; set; }

        [CsvHelper.Configuration.Attributes.Name("开票日期")]
        [JsonProperty("开票日期")]
        public string IssueDate { get; set; }

        [CsvHelper.Configuration.Attributes.Name("项目名称")]
        [JsonProperty("项目名称")]
        public string ItemName { get; set; }

        [CsvHelper.Configuration.Attributes.Name("数量")]
        [JsonProperty("数量")]
        public string Quantity { get; set; }

        [CsvHelper.Configuration.Attributes.Name("金额")]
        [JsonProperty("金额")]
        public string Amount { get; set; }

        [CsvHelper.Configuration.Attributes.Name("税率")]
        [JsonProperty("税率")]
        public string TaxRate { get; set; }

        [CsvHelper.Configuration.Attributes.Name("税额")]
        [JsonProperty("税额")]
        public string TaxAmount { get; set; }

        [CsvHelper.Configuration.Attributes.Name("价税合计")]
        [JsonProperty("价税合计")]
        public string TotalAmount { get; set; }

        [CsvHelper.Configuration.Attributes.Name("发票票种")]
        [JsonProperty("发票票种")]
        public string InvoiceType { get; set; }

        [CsvHelper.Configuration.Attributes.Name("类别")]
        [JsonProperty("类别")]
        public string Category { get; set; }
    }
}
