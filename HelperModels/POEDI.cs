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
    public class POEDI 
    {

        public partial class PO
        {
            public static PO FromJson(string json) => JsonConvert.DeserializeObject<PO>(json, POEDI.Converter.Settings);
        }

        //public static class Serialize  // NOT NEEDED 
        //{
        //    public static string ToJson(this PO  self) => JsonConvert.SerializeObject(self, Converter.Settings);
        //    //return JsonConvert.SerializeObject(self, Converter.Settings);
        //}

        public partial class POOrder
        {
            [JsonProperty("Order")]
            public PO order { get; set; }
        }

        public partial class PO
        {
            [JsonProperty("Header")]
            public Header header { get; set; }

           //[JsonProperty("placeholder")]
            //public string PlaceHolder { get; set; }

            [JsonProperty("LineItem")]
            public List<LineItem> LineItem { get; set; }


            [JsonProperty("Summary")]
            public OrderSummary ordersummary { get; set; }
           
        }

        public class Header
        {
            [JsonProperty("OrderHeader")]
            public OrderHeader OrderHeader { get; set; }

            [JsonProperty("Dates")]
            public List<DateInfo> Date { get; set; }

            //[JsonProperty("Address")]
            //public Address address { get; set; }

            [JsonProperty("Address")]
            public List<Address> address { get; set; }

            //optional 
            [JsonProperty("FOBRelatedInstruction")]
            public List<FOBRelatedInstruction> FOBRelatedInstruction { get; set; }

            //Optional 
            [JsonProperty("Notes")]
            public List<Notes> notes { get; set; }

            //[JsonProperty("Terms")]
           // public Terms  Terms { get; set; }

            [JsonProperty("Terms")]
            public List<Terms> Terms { get; set; }
        }


        //Header Sub-Class
        public class OrderHeader
        {
            [JsonProperty("PurchaseOrderNumber")]
            public string PurchaseOrderNumber { get; set; }

            [JsonProperty("TsetPurposeCode")]
            public string TsetPurposeCode { get; set; }

            [JsonProperty("PrimaryPOTypeCode")]
            public string PrimaryPOTypeCode { get; set; }

            [JsonProperty("PurchaseOrderDate")]
            public DateTime PurchaseOrderDate { get; set; }

            [JsonProperty("Vendor")]
            public string Vendor { get; set; }

            [JsonProperty("CustomerOrderNumber")]
            public string CustomerOrderNumber { get; set; }
        }

        public class DateInfo
        {
            //[JsonProperty("POStartDate")]
            //public DateTime POStartDate { get; set; }

            //[JsonProperty("POCancelDate")]
            //public DateTime POCancelDate { get; set; }

            [JsonProperty("DateTimeQualifier")]
            public string Datetimequalifier {get; set;}


            [JsonProperty("Date")]
            public DateTime Date   { get; set; }

        }

        public class Address
        {
            [JsonProperty("AddressTypeCode")]
            public string AddressTypeCode { get; set; }

            [JsonProperty("LocationCodeQualifier")]
            public string LocationCodeQualifier { get; set; }

            [JsonProperty("AddressLocationNumber")]
            public string AddressLocationNumber { get; set; }

            [JsonProperty("AddressName")]
            public string AddressName { get; set; }

            [JsonProperty("Address1")]
            public string Address1 { get; set; }

            [JsonProperty("City")]
            public string City { get; set; }

            [JsonProperty("State")]
            public string State { get; set; }

            [JsonProperty("PostalCode")]
            public string PostalCode { get; set; }

            [JsonProperty("Country")]
            public string Country { get; set; }
        }

        public class Notes
        {
            [JsonProperty("NoteCode")]
            public string NoteCode { get; set; }

            [JsonProperty("Note")]
            public string Note { get; set; }
        }
         
        //Optional 
        public class FOBRelatedInstruction
        {
            [JsonProperty("FOBPayCode")]
            public string FOBPayCode { get; set; }

            [JsonProperty("FOBLocationQualifier")]
            public string FOBLocationQualifier { get; set; }

            [JsonProperty("FOBLocationDescription")]
            public string FOBLocationDescription { get; set; }

            [JsonProperty("FOBTitlePassageCode")]
            public string FOBTitlePassageCode { get; set; }

            [JsonProperty("FOBTitlePassageLocation")]
            public string FOBTitlePassageLocation { get; set; }

            [JsonProperty("TransportationTermsType")]
            public string TransportationTermsType { get; set; }

            [JsonProperty("TransportationTerms")]
            public string TransportationTerms { get; set; }

            [JsonProperty("RiskOfLossCode")]
            public string RiskOfLossCode { get; set; }

            [JsonProperty("Description")]
            public string Description { get; set; }
        }

        //optional
        public class Contact
        {
            [JsonProperty("ContactTypeCode")]
            public string ContactTypeCode { get; set; }

            [JsonProperty("ContactName")]
            public string ContactName { get; set; }

            [JsonProperty("PrimaryPhone")]
            public string PrimaryPhone { get; set; }

            [JsonProperty("PrimaryEmail")]
            public string PrimaryEmail { get; set; }
}

        public class Terms   
        {
            [JsonProperty("TermsDescription")]
            public string termsDescription { get; set; }
        }

            public class LineItem
        {
            [JsonProperty("OrderLine")]
            public Orderline orderline { get; set; }

            [JsonProperty("ProductOrItemDescription")]
            public List<ProductorItemDescriptions> productorItemDescriptions { get; set; }
        }

        public class Orderline
        {
            [JsonProperty("LineSequenceNumber")]
            public string LineSequenceNumber { get; set; }

            [JsonProperty("VendorPartNumber")]
            public string VendorPartNumber { get; set; }

            [JsonProperty("ConsumerPackageCode")]
            public string ConsumerPackageCode { get; set; }

            [JsonProperty("OrderQty")]
            public string OrderQty { get; set; }

            [JsonProperty("OrderQtyUOM")]
            public string OrderQtyUOM { get; set; }

            [JsonProperty("PurchasePrice")]
            public string PurchasePrice { get; set; }

            [JsonProperty("ExtendedItemTotal")]
            public string ExtendedItemTotal { get; set; }

            [JsonProperty("ProductColorCode")]
            public string Color { get; set; }

            [JsonProperty("ProductSizeCode")]
            public string Size { get; set; }
        }

        public class ProductorItemDescriptions
        {
            [JsonProperty("ProductCharacteristicCode")]
            public string ProductCharacteristicCode { get; set; }

            [JsonProperty("ProductDescription")]
            public string ProductDescription { get; set; }
        }

        public class OrderSummary
        {
            [JsonProperty("TotalAmount")]
            public string TotalAmount { get; set; }

            [JsonProperty("TotalQuantity")]
            public string TotalQuantity { get; set; }

            [JsonProperty("TotalLineItemNumber")]
            public string TotalLineItemNumber { get; set; }
        }


        public class Items
        {
            [JsonProperty("Items")]
            public List<Item> items { get; set; }
        }

        public class Item
        {
            [JsonProperty("LineNumber")]
            public string LineNumber { get; set; }

            [JsonProperty("VendorpartNumber")]
            public string VendorpartNumber { get; set; }

            [JsonProperty("UPC")]
            public string UPC { get; set; }

            [JsonProperty("Quantity")]
            public string Quantity { get; set; }

            [JsonProperty("UOM")]
            public string UOM { get; set; }

            [JsonProperty("UnitPrice")]
            public string UnitPrice { get; set; }
        }

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
    }

    public class POOutput
    {
        public string LineNumber { get; set; }
        public string Manufacturer { get; set; }
        public string PONumber { get; set; }
        public string Posted { get; set; }
        public string Style { get; set; }
        public string Description { get; set; }
        public string Color { get; set; }
        public string ColorCode { get; set; }
        public string Size { get; set; }
        public string UPC { get; set; }
        public string QtyOrdered { get; set; }
        public string Cost { get; set; }
        public string ExtCost { get; set; }
        public string Receive { get; set; }
        public string ColorOverride { get; set; }
        public string StyleOverride { get; set; }
        public string Closeout { get; set; }
        public string SKUCloseout { get; set; }
        public string Backorder { get; set; }

    }

}

