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
    class SPS856
    {

        //public partial class SHFile
        //{
        //    public static File856 FromJson(string json) => JsonConvert.DeserializeObject<File856>(json, SPS856.Converter.Settings);
        //}

            /*
        internal static class Converter
        {
            public static readonly JsonSerializerSettings Settings = new JsonSerializerSettings
            {
                MetadataPropertyHandling = MetadataPropertyHandling.Ignore,
                NullValueHandling = NullValueHandling.Ignore,
                MissingMemberHandling = MissingMemberHandling.Ignore,
                DateParseHandling = DateParseHandling.None,
                Converters =
            {
                new IsoDateTimeConverter { DateTimeStyles = DateTimeStyles.AssumeUniversal }
            },
            };
        }
        */

        public partial class SHFile
        {
            public class File856
            {
                [JsonProperty("Header")]
                public Header Header { get; set; }

                [JsonProperty("OrderLevel")]
                public List<OrderLevel> OrderLevel { get; set; }

                [JsonProperty("Summary")]
                public Summary Summary { get; set; }
            }
        }

        public class Header
        {
            [JsonProperty("ShipmentHeader")]
            public ShipmentHeader ShipmentHeader { get; set; }

            [JsonProperty("Address")]
            public List<Address> Address { get; set; }

            [JsonProperty("CarrierInformation")]
            public List<CarrierInformation> CarrierInformation { get; set; }

            [JsonProperty("QuantityAndWeight")]
            public List<QuantityAndWeight> QuantityAndWeight { get; set; }

        }

        public class OrderLevel
        {
            [JsonProperty("OrderHeader")]
            public OrderHeader OrderHeader { get; set; }

            [JsonProperty("PackLevel")]
            public List<PackLevel> PackLevel { get; set; }
    }

        public class Summary
        {
            [JsonProperty("TotalOrders")]
            public int TotalOrders { get; set; }

            [JsonProperty("TotalLineItemNumber")]
            public int TotalLineItemNumber { get; set; }

        }

        ///////////////////////////////////////////
        /// Header

        public class ShipmentHeader
        {
            [JsonProperty("ShipmentIdentification")]
            public string ShipmentIdentification { get; set; }

            [JsonProperty("ShipDate")]
            public string ShipDate { get; set; }

            [JsonProperty("TsetPurposeCode")]
            public string TsetPurposeCode { get; set; }

            [JsonProperty("ShipNoticeDate")]
            public string ShipNoticeDate { get; set; }

            [JsonProperty("BillOfLadingNumber")]
            public string BillOfLadingNumber { get; set; }

			[JsonProperty("CurrentScheduledDeliveryDate")]
            public string CurrentScheduledDeliveryDate { get; set; }

            [JsonProperty("CarrierProNumber")]
            public string CarrierProNumber { get; set; }
}

        public class Address
        {
            [JsonProperty("AddressTypeCode")]
            public string AddressTypeCode { get; set; }

            [JsonProperty("AddressName")]
            public string AddressName { get; set; }

            [JsonProperty("Address1")]
            public string Address1 { get; set; }

            //Address2 { get; set; }

            [JsonProperty("City")]
            public string City { get; set; }

            [JsonProperty("PostalCode")]
            public string PostalCode { get; set; }
        }

        public class CarrierInformation
        {
            [JsonProperty("CarrierAlphaCode")]
            public string CarrierAlphaCode { get; set; }

            [JsonProperty("CarrierRouting")]
            public string CarrierRouting { get; set; }
        }

        public class QuantityAndWeight
        {
            [JsonProperty("PackingMedium")]
            public string PackingMedium { get; set; }

            [JsonProperty("LadingQuantity")]
            public int LadingQuantity { get; set; }

            [JsonProperty("WeightQualifier")]
            public string WeightQualifier { get; set; }
        }

        ///////////////////////////////////////
        /// OrderLevel

        public class OrderHeader
        {
            [JsonProperty("PurchaseOrderNumber")]
            public string PurchaseOrderNumber { get; set; }

            [JsonProperty("PurchaseOrderDate")]
            public string PurchaseOrderDate { get; set; }

            [JsonProperty("Vendor")]
            public string Vendor { get; set; }
        }

        public class PackLevel
        {
            [JsonProperty("Pack")]
            public Pack Pack { get; set; }

            [JsonProperty("ItemLevel")]
            public List<ItemLevel> ItemLevel { get; set; }
    }

        //////////////////////////////////////////
        /// PackLevel

        public class Pack
        {
            [JsonProperty("PackLevelType")]
            public string PackLevelType { get; set; }

            [JsonProperty("ShippingSerialID")]
            public string ShippingSerialID { get; set; }
        }

        public class ItemLevel
        {
            [JsonProperty("ShipmentLine")]
            public ShipmentLine ShipmentLine { get; set; }

            [JsonProperty("ProductOrItemDescription")]
            public List<ProductOrItemDescription> ProductOrItemDescription { get; set; } 

        }

        /// ////////////////////////////////////////////////////////
        /// ItemLevel
        public class ShipmentLine
        {
            [JsonProperty("LineSequenceNumber")]
            public string LineSequenceNumber { get; set; }

            [JsonProperty("VendorPartNumber")]
            public string VendorPartNumber { get; set; }

            [JsonProperty("ConsumerPackageCode")]
            public string ConsumerPackageCode { get; set; }

            [JsonProperty("ShipQty")]
            public decimal	ShipQty { get; set; }

            [JsonProperty("ShipQtyUOM")]
            public string ShipQtyUOM { get; set; }

            [JsonProperty("ProductSizeDescription")]
            public string ProductSizeDescription { get; set; }

            [JsonProperty("ProductColorDescription")]
            public string ProductColorDescription { get; set; }

        }

        public class ProductOrItemDescription
        {
            [JsonProperty("ProductCharacteristicCode")]
            public string ProductCharacteristicCode { get; set; }

            [JsonProperty("ProductDescription")]
            public string ProductDescription { get; set; }
        }
    }
}
