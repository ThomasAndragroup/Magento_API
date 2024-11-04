using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Globalization;
using Newtonsoft.Json.Converters;

namespace MagentoProductAPI.HelperModels
{
    class SP846
    {

        public class File846
        {
            public partial class ItemRegistry
            {
                
                [JsonProperty("type")]
                public string type { get; set; }

                [JsonProperty("additionalProperties")]
                public string additionalProperties { get; set; }

                [JsonProperty("required")]
                public List<string> required { get; set; }
                
            
                [JsonProperty("Header")]
                public Header Header { get; set; }

                [JsonProperty("Structure")]
                public Structure Structure { get; set; }

                [JsonProperty("Summary")]
                public Summary Summary { get; set; }
            }
        }

        public class Header
        {
            [JsonProperty("HeaderReport")]
            public HeaderReport HeaderReport { get; set; }

            [JsonProperty("Address")]
            public Address Address { get; set; }

        }

        public class HeaderReport
        {
            [JsonProperty("DocumentId")]
            public string DocumentId { get; set; }

            [JsonProperty("TsetPurposeCode")]
            public string TsetPurposeCode { get; set; }

            [JsonProperty("ReportTypeCode")]
            public string ReportTypeCode { get; set; }

            [JsonProperty("InventoryDate")]
            public DateTime InventoryDate { get; set; }

            [JsonProperty("InventoryTime")]
            public string InventoryTime { get; set; }

            [JsonProperty("Vendor")]
            public string Vendor { get; set; }
        }

        public class Address
        {
            [JsonProperty("AddressTypeCode")]
            public string AddressTypeCode { get; set; }

            [JsonProperty("AddressName")]
            public string AddressName { get; set; }

            [JsonProperty("Address1")]
            public string Address1 { get; set; }

            [JsonProperty("Address2")]
            public string Address2 { get; set; }

            [JsonProperty("Address3")]
            public string Address3 { get; set; }

            [JsonProperty("City")]
            public string City { get; set; }

            [JsonProperty("State")]
            public string State { get; set; }

            [JsonProperty("PostalCode")]
            public string PostalCode { get; set; }

            [JsonProperty("Country")]
            public string Country { get; set; }
        }

        /// <summary>
        /// //////////////////////////////////////////////////////////////////////////////
        /// </summary>
        public class Structure
        {
            [JsonProperty("LineItem")]
            public LineItem LineItem { get; set; }
        }

        public class LineItem
        {
            [JsonProperty("Dates")]
            public InventoryLine Dates { get; set; }

            [JsonProperty("Datesxx")]
            public Date Datesxx { get; set; }

            [JsonProperty("PriceDetails")]
            public List<PriceDetail> PriceDetails { get; set; }

            [JsonProperty("ProductOrItemDescription")]
            public ProductOrItemDescription ProductOrItemDescription { get; set; }

            [JsonProperty("QuantitiesSchedulesLocations")]
            public List<QuantitiesSchedulesLocations> QuantitiesSchedulesLocations { get; set; }

        }

        public class InventoryLine
        {
            [JsonProperty("LineSequenceNumber")]
            public int LineSequenceNumber { get; set; }

            [JsonProperty("VendorPartNumber")]
            public int VendorPartNumber { get; set; }

            [JsonProperty("ConsumerPackageCode")]
            public int ConsumerPackageCode { get; set; }

            [JsonProperty("ProductSizeDescription")]
            public string ProductSizeDescription { get; set; }

            [JsonProperty("ProductColorDescription")]
            public string ProductColorDescription { get; set; }
        }

        public class Date
        {
            [JsonProperty("DateTimeQualifier")]
            public int DateTimeQualifier { get; set; }

            [JsonProperty("Date")]
            public DateTime Datexx { get; set; }
        }

        public class PriceDetail
        {
            [JsonProperty("PriceTypeIDCode")]
            public string PriceTypeIDCode { get; set; }

            [JsonProperty("UnitPrice")]
            public decimal UnitPrice { get; set; }
        }

        public class ProductOrItemDescription
        {
            [JsonProperty("ProductCharacteristicCode")]
            public string ProductCharacteristicCode { get; set; }

            [JsonProperty("ProductDescription")]
            public string ProductDescription { get; set; }
        }

        public class QuantitiesSchedulesLocations
        {
            [JsonProperty("QuantityQualifier")]
            public int QuantityQualifier { get; set; }

            [JsonProperty("TotalQty")]
            public int TotalQty { get; set; }

            [JsonProperty("TotalQtyUOM")]
            public string TotalQtyUOM { get; set; }

            [JsonProperty("Dates")]
            public List<Date> Dates { get; set; }
        }


        /// <summary>
        /// ////////////////////////////////////////////////////////////////////////////
        /// </summary>
        public class Summary
        {
            [JsonProperty("TotalLineItemNumber")]
            public int TotalLineItemNumber { get; set; }
        }
    }
    //}
    //}

    //May not need this:
    class SP846_File
    {

        [JsonProperty("FileInfo")]
        public string FileInfo  { get; set; }

        [JsonProperty("InventoryDate")]
        public string InventoryDate { get; set; }

        [JsonProperty("sku")]
        public string sku { get; set; }

        [JsonProperty("TotalQty")]
        public int TotalQty { get; set; }

        [JsonProperty("TotalQtyUOM")]
        public int TotalQtyUOM { get; set; }

        [JsonProperty("LineItemDate")]
        public string LineItemDate { get; set; }

        [JsonProperty("DocumentId")]
        public string DocumentId { get; set; }

        [JsonProperty("AddressInfo")]
        public string AddressInfo { get; set; }

        [JsonProperty("UnitPrice")]
        public decimal UnitPrice { get; set; }

        [JsonProperty("LineSequenceNumber")]
        public int LineSequenceNumber { get; set; }

        [JsonProperty("ProductDescription")]
        public string ProductDescription { get; set; }

        [JsonProperty("VendorPartNumber")]
        public string VendorPartNumber { get; set; }

        [JsonProperty("ConsumerPackageCode")]
        public string ConsumerPackageCode { get; set; }

        [JsonProperty("ProductSizeDescription")]
        public string ProductSizeDescription { get; set; }

        [JsonProperty("ProductColorDescription")]
        public string ProductColorDescription { get; set; }
    }
}
