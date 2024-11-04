using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using Newtonsoft.Json;
using Magento_MCP;
using System.Net.Http;
using System.Configuration;
using System.IO;
using System.Text;

namespace MagentoProductAPI
{
    class Klaviyo
    {
        /*
         feeds:
            SisterHoodBra - DONE
            SisterHoodPanties - DONE

            fulfilled his -1 0  1/day
            fulfilled her -1 0  1/day

            phonenumberupdate   2/day
            
            placedmin his -30 x
            cancelledmin his -30 x
            returnedmin his -30x

            placedmin her -30 x
            cancelledmin her -30 x
            returnedmin her -30x

            Case "placedmin"
                sql = String.Format("SELECT orderno FROM [hercust].[dbo].[orders] WHERE store = '{0}' and amazon = 0 and retailerid = 0 and not confirmemail like '%@orders.fiftyone.com' and TimeStamp >= dateadd(mi, {1}, GetDate()) and TimeStamp < GetDate() order by orderno", store, fromDate, endDate)
            Case "fulfilledmin"
                sql = String.Format("SELECT orderno FROM [hercust].[dbo].[orders] WHERE store = '{0}' and shipped = 1 and amazon = 0 and retailerid = 0 and not confirmemail like '%@orders.fiftyone.com' and ShipDate >= dateadd(mi, {1}, GetDate()) and ShipDate < GetDate() order by orderno", store, fromDate, endDate)
            Case "cancelledmin"
                sql = String.Format("SELECT orderno FROM [hercust].[dbo].[orders] WHERE store = '{0}' and cancel = 1 and amazon = 0 and retailerid = 0 and not confirmemail like '%@orders.fiftyone.com' and CancelledWhen >= dateadd(mi, {1}, GetDate()) and CancelledWhen < GetDate() order by orderno ", store, fromDate, endDate)
            Case "returnedmin"
                sql = String.Format("SELECT DISTINCT OIR.ORDERNO FROM hercust..OrderItemReturns OIR INNER JOIN hercust..orders OO ON OO.OrderNo = OIR.OrderNo AND OO.Returned > 0 AND Store = '{0}' WHERE OIR.ReturnDate BETWEEN dateadd(mi, {1}, GetDate()) AND GetDate() ORDER BY 1", store, fromDate, endDate)



    */

        public class SisterhoodBra
        {
            [JsonProperty("data")]
            public SisterhoodBraData data { get; set; }

        }

        public class SisterhoodBraData
        {
            [JsonProperty("type")]
            public string type { get; set; }

            [JsonProperty("attributes")]
            public SisterhoodBraAttributes attributes { get; set; }

        }

        public class SisterhoodBraAttributes
        {
            [JsonProperty("email")]
            public string email { get; set; }

            [JsonProperty("first_name")]
            public string first_name { get; set; }

            [JsonProperty("last_name")]
            public string last_name { get; set; }

            [JsonProperty("company_id")]
            public string company_id { get; set; }

            [JsonProperty("properties")]
            public SisterhoodBraProperties properties { get; set; }
        }

        public class SisterhoodBraProperties
        {
            [JsonProperty("SisterHoodBra")]
            public string SisterHoodBra { get; set; }

            [JsonProperty("SisterHoodBraImage")]
            public string SisterHoodBraImage { get; set; }

            [JsonProperty("SisterHoodBraName")]
            public string SisterHoodBraName { get; set; }

            [JsonProperty("SisterHoodBraUrl")]
            public string SisterHoodBraUrl { get; set; }

            [JsonProperty("SisterHoodBraSize")]
            public string SisterHoodBraSize { get; set; }

            [JsonProperty("OrderDate")]
            public string OrderDate { get; set; }

            [JsonProperty("product1")]
            public string product1 { get; set; }

            [JsonProperty("product2")]
            public string product2 { get; set; }

            [JsonProperty("product3")]
            public string product3 { get; set; }

            [JsonProperty("bra1image")]
            public string bra1image { get; set; }

            [JsonProperty("bra1name")]
            public string bra1name { get; set; }

            [JsonProperty("bra1url")]
            public string bra1url { get; set; }

            [JsonProperty("bra2image")]
            public string bra2image { get; set; }

            [JsonProperty("bra2name")]
            public string bra2name { get; set; }

            [JsonProperty("bra2url")]
            public string bra2url { get; set; }

            [JsonProperty("bra3image")]
            public string bra3image { get; set; }

            [JsonProperty("bra3name")]
            public string bra3name { get; set; }

            [JsonProperty("bra3url")]
            public string bra3url { get; set; }

            [JsonProperty("SisterHoodBraBrand")]
            public string SisterHoodBraBrand { get; set; }

            [JsonProperty("bra1brand")]
            public string bra1brand { get; set; }

            [JsonProperty("bra2brand")]
            public string bra2brand { get; set; }

            [JsonProperty("bra3brand")]
            public string bra3brand { get; set; }

            [JsonProperty("SisterHoodBraStyle")]
            public string SisterHoodBraStyle { get; set; }

            [JsonProperty("bra1style")]
            public string bra1style { get; set; }

            [JsonProperty("bra2style")]
            public string bra2style { get; set; }

            [JsonProperty("bra3style")]
            public string bra3style { get; set; }

            [JsonProperty("SisterHoodBraTS")]
            public string SisterHoodBraTS { get; set; }
        }



        /// ////////////////////////////////////////////////////
        public class SisterhoodPanty
        {
            [JsonProperty("data")]
            public SisterhoodPantyData data { get; set; }

        }

        public class SisterhoodPantyData
        {
            [JsonProperty("type")]
            public string type { get; set; }

            [JsonProperty("attributes")]
            public SisterhoodPantyAttributes attributes { get; set; }

        }

        public class SisterhoodPantyAttributes
        {
            [JsonProperty("email")]
            public string email { get; set; }

            [JsonProperty("first_name")]
            public string first_name { get; set; }

            [JsonProperty("last_name")]
            public string last_name { get; set; }

            [JsonProperty("company_id")]
            public string company_id { get; set; }

            [JsonProperty("properties")]
            public SisterhoodPantyProperties properties { get; set; }
        }

        public class SisterhoodPantyProperties
        {
            [JsonProperty("SisterHoodPanty")]
            public string SisterHoodPanty { get; set; }

            [JsonProperty("SisterHoodPantyImage")]
            public string SisterHoodPantyImage { get; set; }

            [JsonProperty("SisterHoodPantyName")]
            public string SisterHoodPantyName { get; set; }

            [JsonProperty("SisterHoodPantyUrl")]
            public string SisterHoodPantyUrl { get; set; }

            [JsonProperty("SisterHoodPantySize")]
            public string SisterHoodPantySize { get; set; }

            [JsonProperty("OrderDate")]
            public string OrderDate { get; set; }

            [JsonProperty("panty1")]
            public string panty1 { get; set; }

            [JsonProperty("panty2")]
            public string panty2 { get; set; }

            [JsonProperty("panty3")]
            public string panty3 { get; set; }

            [JsonProperty("pantyimage")]
            public string panty1image { get; set; }

            [JsonProperty("panty1name")]
            public string panty1name { get; set; }

            [JsonProperty("panty1url")]
            public string panty1url { get; set; }

            [JsonProperty("panty2image")]
            public string panty2image { get; set; }

            [JsonProperty("panty2name")]
            public string panty2name { get; set; }

            [JsonProperty("panty2url")]
            public string panty2url { get; set; }

            [JsonProperty("panty3image")]
            public string panty3image { get; set; }

            [JsonProperty("panty3name")]
            public string panty3name { get; set; }

            [JsonProperty("panty3url")]
            public string panty3url { get; set; }

            [JsonProperty("SisterHoodPantyBrand")]
            public string SisterHoodPantyBrand { get; set; }

            [JsonProperty("panty1brand")]
            public string panty1brand { get; set; }

            [JsonProperty("panty2brand")]
            public string panty2brand { get; set; }

            [JsonProperty("panty3brand")]
            public string panty3brand { get; set; }

            [JsonProperty("SisterHoodPantyStyle")]
            public string SisterHoodPantyStyle { get; set; }

            [JsonProperty("panty1style")]
            public string panty1style { get; set; }

            [JsonProperty("panty2style")]
            public string panty2style { get; set; }

            [JsonProperty("panty3style")]
            public string panty3style { get; set; }

            [JsonProperty("SisterHoodPantyTS")]
            public string SisterHoodPantyTS { get; set; }
        }


        public static Boolean SisterhoodBras_API()
        {
            Boolean Returnvalue = true;
            DataTable dt;
            SisterhoodBra SHB;
            String APIResult;


            dt = Helper.Sql_Misc_Fetch("EXEC herroom..[proc_mag_get_sisterhoodbra_klaviyoapi]");
            foreach (DataRow dr in dt.Rows)
            {
                SHB = new SisterhoodBra();
                SHB.data = new SisterhoodBraData();

                SHB.data.type = "profile";

                SHB.data.attributes = new SisterhoodBraAttributes();
                SHB.data.attributes.properties = new SisterhoodBraProperties();

                SHB.data.attributes.email = dr["confirmemail"].ToString();
                SHB.data.attributes.first_name = dr["FirstName"].ToString();
                SHB.data.attributes.last_name = dr["LastName"].ToString();
                SHB.data.attributes.company_id = "TUJuFB";      // harded code now, may not need

                SHB.data.attributes.properties.SisterHoodBra = dr["StyleNumber"].ToString();
                SHB.data.attributes.properties.SisterHoodBraImage = dr["Style0Image"].ToString();
                SHB.data.attributes.properties.SisterHoodBraName = dr["Style0Name"].ToString();
                SHB.data.attributes.properties.SisterHoodBraUrl = dr["Style0URL"].ToString();
                SHB.data.attributes.properties.SisterHoodBraSize = dr["Size"].ToString();
                SHB.data.attributes.properties.OrderDate = dr["OrderDate"].ToString();
                SHB.data.attributes.properties.product1 = dr["StyleNo1"].ToString();
                SHB.data.attributes.properties.product2 = dr["StyleNo2"].ToString();
                SHB.data.attributes.properties.product3 = dr["StyleNo3"].ToString();
                SHB.data.attributes.properties.bra1image = dr["Style1Image"].ToString();
                SHB.data.attributes.properties.bra1name = dr["Style1Name"].ToString();
                SHB.data.attributes.properties.bra1url = dr["Style1URL"].ToString();
                SHB.data.attributes.properties.bra2image = dr["Style2Image"].ToString();
                SHB.data.attributes.properties.bra2name = dr["Style2Name"].ToString();
                SHB.data.attributes.properties.bra2url = dr["Style2URL"].ToString();
                SHB.data.attributes.properties.bra3image = dr["Style3Image"].ToString();
                SHB.data.attributes.properties.bra3name = dr["Style3Name"].ToString();
                SHB.data.attributes.properties.bra3url = dr["Style3URL"].ToString();
                SHB.data.attributes.properties.SisterHoodBraBrand = dr["SisterHoodBraBrand"].ToString();
                SHB.data.attributes.properties.bra1brand = dr["bra1brand"].ToString();
                SHB.data.attributes.properties.bra2brand = dr["bra2brand"].ToString();
                SHB.data.attributes.properties.bra3brand = dr["bra3brand"].ToString();
                SHB.data.attributes.properties.SisterHoodBraStyle = dr["SisterHoodBraStyle"].ToString();
                SHB.data.attributes.properties.bra1style = dr["bra1style"].ToString();
                SHB.data.attributes.properties.bra2style = dr["bra2style"].ToString();
                SHB.data.attributes.properties.bra3style = dr["bra3style"].ToString();
                SHB.data.attributes.properties.SisterHoodBraTS = dr["TS"].ToString();

                var TM_Json = JsonConvert.SerializeObject(SHB, Formatting.Indented, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });

                Console.WriteLine(TM_Json);
                //Console.WriteLine(" ------------------ ");

                APIResult = Helper.MagentoApiPush_Klaviyo("client", TM_Json, "HER");

                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine(dr["confirmemail"].ToString() + " :: " + APIResult);
                    Console.WriteLine(" ------------------ ");
                }
            }

            return Returnvalue;
        }

        public static Boolean SisterhoodPanty_API()
        {
            Boolean Returnvalue = true;
            DataTable dt;
            SisterhoodPanty SHB;
            String APIResult;

            dt = Helper.Sql_Misc_Fetch("EXEC herroom..[proc_mag_get_sisterhoodpanty_klaviyoapi]");
            foreach (DataRow dr in dt.Rows)
            {
                SHB = new SisterhoodPanty();
                SHB.data = new SisterhoodPantyData();

                SHB.data.type = "profile";

                SHB.data.attributes = new SisterhoodPantyAttributes();
                SHB.data.attributes.properties = new SisterhoodPantyProperties();

                SHB.data.attributes.email = dr["confirmemail"].ToString();
                SHB.data.attributes.first_name = dr["FirstName"].ToString();
                SHB.data.attributes.last_name = dr["LastName"].ToString();
                SHB.data.attributes.company_id = "XhueMt";      // harded code now, may not need

                SHB.data.attributes.properties.SisterHoodPanty = dr["StyleNumber"].ToString();
                SHB.data.attributes.properties.SisterHoodPantyImage = dr["Style0Image"].ToString();
                SHB.data.attributes.properties.SisterHoodPantyName = dr["Style0Name"].ToString();
                SHB.data.attributes.properties.SisterHoodPantyUrl = dr["Style0URL"].ToString();
                SHB.data.attributes.properties.SisterHoodPantySize = dr["Size"].ToString();

                SHB.data.attributes.properties.OrderDate = dr["OrderDate"].ToString();
                SHB.data.attributes.properties.panty1 = dr["StyleNo1"].ToString();
                SHB.data.attributes.properties.panty2 = dr["StyleNo2"].ToString();
                SHB.data.attributes.properties.panty3 = dr["StyleNo3"].ToString();
                SHB.data.attributes.properties.panty1image = dr["Style1Image"].ToString();
                SHB.data.attributes.properties.panty1name = dr["Style1Name"].ToString();
                SHB.data.attributes.properties.panty1url = dr["Style1URL"].ToString();
                SHB.data.attributes.properties.panty2image = dr["Style2Image"].ToString();
                SHB.data.attributes.properties.panty2name = dr["Style2Name"].ToString();
                SHB.data.attributes.properties.panty2url = dr["Style2URL"].ToString();
                SHB.data.attributes.properties.panty3image = dr["Style3Image"].ToString();
                SHB.data.attributes.properties.panty3name = dr["Style3Name"].ToString();
                SHB.data.attributes.properties.panty3url = dr["Style3URL"].ToString();
                SHB.data.attributes.properties.SisterHoodPantyBrand = dr["SisterHoodpantyBrand"].ToString();
                SHB.data.attributes.properties.panty1brand = dr["panty1brand"].ToString();
                SHB.data.attributes.properties.panty2brand = dr["panty2brand"].ToString();
                SHB.data.attributes.properties.panty3brand = dr["panty3brand"].ToString();
                SHB.data.attributes.properties.SisterHoodPantyStyle = dr["SisterHoodPantyStyle"].ToString();
                SHB.data.attributes.properties.panty1style = dr["panty1style"].ToString();
                SHB.data.attributes.properties.panty2style = dr["panty2style"].ToString();
                SHB.data.attributes.properties.panty3style = dr["panty3style"].ToString();
                SHB.data.attributes.properties.SisterHoodPantyTS = dr["TS"].ToString();

                var TM_Json = JsonConvert.SerializeObject(SHB, Formatting.Indented, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });

                Console.WriteLine(TM_Json);
                //Console.WriteLine(" ------------------ ");

                APIResult = Helper.MagentoApiPush_Klaviyo("client", TM_Json, "HER");
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine(dr["confirmemail"].ToString() + " :: " + APIResult);
                    Console.WriteLine(" ------------------ ");
                }
            }

            return Returnvalue;
        }


        /// ///////////////////////////////////////////////////////////////////////////////////

        private class PhoneNumberUpdate
        {
            [JsonProperty("data")]
            public PhoneNumberUpdateData data { get; set; }
        }

        private class PhoneNumberUpdateData
        {
            [JsonProperty("attributes")]
            public PhoneNumberUpdateAttributes attributes { get; set; }
        }

        private class PhoneNumberUpdateAttributes
        {
            [JsonProperty("email")]
            public string email { get; set; }

            [JsonProperty("phone_number")]
            public string phone_number { get; set; }

            [JsonProperty("company_id")]
            public string company_id { get; set; }

        }

        //TEST
        public static Boolean PhoneNumberUpdate_API()
        {
            Boolean Returnvalue = true;
            DataTable dt;
            PhoneNumberUpdate PN;
            String APIResult;

            dt = Helper.Sql_Misc_Fetch("EXEC Tempwork..[proc_cust_klaviyo_updates]");
            foreach (DataRow dr in dt.Rows)
            {
                PN = new PhoneNumberUpdate();
                PN.data = new PhoneNumberUpdateData();

                PN.data.attributes.email = dr["email"].ToString();
                PN.data.attributes.phone_number = "+1" + dr["phonenumber"].ToString();
               // PN.data.attributes.company_id = "XhueMt";

                var TM_Json = JsonConvert.SerializeObject(PN, Formatting.Indented, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });

                APIResult = Helper.MagentoApiPush_Klaviyo("client", TM_Json, "HER");  // HER ?? 

                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine(TM_Json);
                    Console.WriteLine(APIResult);
                    Console.WriteLine(" ------------------ ");
                }

                
               

            }
            //"EXEC Tempwork..proc_cust_klaviyo_updates"
            /*
                custno	email	                    phonenumber
                2050356	Sloughlin@tampabay.rr.com	8133613395
                3755677	jenwales@comcast.net	    7703956007

             {
                "data": {
                    "type": "profile",
                    "attributes": {
                        "email": "thomas.tribble@gmail.com",
                        "phone_number": "+12149669587",
                        "company_id": "XhueMt"
                    }
                }
            }
            */

            return Returnvalue;
        }

        /// ///////////////////////////////////////////////////////////////////////////////////////////////
        // ORDER CLASSES

        public class KlaviyoOrder
        {
            [JsonProperty("data")]
            public OrderData data { get; set; }

        }

        public class OrderData
        {

            [JsonProperty("type")]
            public string Type { get; set; }

            [JsonProperty("attributes")]
            public OrderAttributes OrderAttributes { get; set; }

        }

        public class OrderAttributes
        {
            [JsonProperty("properties")]
            public OrderProperties Properties { get; set; }
            //public OrderAttributeProperties properties { get; set; }


            [JsonProperty("time")]
            public string Time { get; set; }

            [JsonProperty("value")]
            public decimal Value { get; set; }

            [JsonProperty("value_currency")]
            public string Value_currency { get; set; }

            [JsonProperty("metric")]
            public OrderAttributeMetric Metric { get; set; }

            [JsonProperty("profile")]
            public OrderProfile profile { get; set; }
        }

        /*
        public class OrderAttributeProperties
        {
            [JsonProperty("orderno")]
            public string Orderno { get; set; }
        }
        */

        public class OrderAttributeMetric
        {
            [JsonProperty("data")]
            public OrderAttributeMetricData Data { get; set; }
        }

        public class OrderAttributeMetricData
        {
            [JsonProperty("type")]
            public string Type { get; set; }

            [JsonProperty("attributes")]
            public OrderAttributeMetricDataAttributes Attributes { get; set; }
        }

        public class OrderAttributeMetricDataAttributes
        {
            [JsonProperty("name")]
            public string Name { get; set; }
        }

        public class OrderProfile
        {
            [JsonProperty("data")]
            public OrderProfileData Data { get; set; }
        }

        public class OrderProfileData
        {
            [JsonProperty("type")]
            public string Type { get; set; }

            [JsonProperty("attributes")]
            public OrderProfileDataAttributes Attributes { get; set; }

            //[JsonProperty("email")]
            //public string email { get; set; }

        }

        public class OrderProfileDataAttributes
        {
            [JsonProperty("properties")]
            public OrderProfileDataAttributesProperties Properties { get; set; }

            [JsonProperty("email")]
            public string Email { get; set; }

        }

        public class OrderProfileDataAttributesProperties
        {
            [JsonProperty("newKey")]
            public string newKey { get; set; }

        }

        /*
        public class Metric
        {
            MetricData
        }

        public class MetricData
        {
            Type

MetricDataAttribute
        }

        public class MetricDataAttribute
        {
            name
        }
        */



        /// ////////////////////////////////////////////////////////////
        public class OrderProperties 
        {
            [JsonProperty("orderno")]
            public string Orderno { get; set; }

            //[JsonProperty("value")]
            //public decimal value { get; set; }

            //[JsonProperty("value_currency")]
            //public string Value_currency { get; set; }


            [JsonProperty("Categories")]
            public List<string> Categories { get; set; }

            [JsonProperty("Subcategories")]
            public List<string> Subcategories { get; set; }

            [JsonProperty("Tastes")]
            public List<string> Tastes { get; set; }

            [JsonProperty("Onsale")]
            public string Onsale { get; set; }

            [JsonProperty("Braddplus")]
            public string Braddplus { get; set; }

            [JsonProperty("Braband")]
            public string Braband { get; set; }

            [JsonProperty("ItemNames")]
            public List<string> ItemNames { get; set; }

            [JsonProperty("Brands")]
            public string Brands { get; set; }

            [JsonProperty("Discount Code")]
            public string DiscountCode { get; set; }

            [JsonProperty("Discount Value")]
            public decimal DiscountValue { get; set; }

            [JsonProperty("Discount Type")]
            public string DiscountType { get; set; }

            [JsonProperty("Items")]
            public List<OrderItem> Items { get; set; }

            //[JsonProperty("billing_address")]
            //public OrderAddress billing_address { get; set; }

           //[JsonProperty("shipping_address")]
            //public OrderAddress shipping_address { get; set; }
        }

        public class OrderItem
        {
            [JsonProperty("ProductID")]
            public string ProductID { get; set; }

            [JsonProperty("SKU")]
            public string SKU { get; set; }

            [JsonProperty("ProductName")]
            public string ProductName { get; set; }

            [JsonProperty("Size")]
            public string Size { get; set; }

            [JsonProperty("Color")]
            public string Color { get; set; }

            [JsonProperty("Quantity")]
            public int Quantity { get; set; }

            [JsonProperty("ItemPrice")]
            public decimal ItemPrice { get; set; }

            [JsonProperty("RowTotal")]
            public decimal RowTotal { get; set; }

            [JsonProperty("ProductURL")]
            public string ProductURL { get; set; }

            [JsonProperty("ImageURL")]
            public string ImageURL { get; set; }

            [JsonProperty("Categories")]
            public List<string> Categories { get; set; }

            [JsonProperty("Brand")]
            public string Brands { get; set; }

            [JsonProperty("onsale")]
            public string onsale { get; set; }

            [JsonProperty("braband")]
            public string braband { get; set; }

            [JsonProperty("braddplus")]
            public string braddplus { get; set; }
                   
        }

        public class OrderAddress
        {
            [JsonProperty("first_name")]
            public string Firstname { get; set; }

            [JsonProperty("last_name")]
            public string Lastname { get; set; }

            [JsonProperty("company")]
            public string Company { get; set; }

            [JsonProperty("address1")]
            public string address1 { get; set; }

            [JsonProperty("address2")]
            public string Address2 { get; set; }

            [JsonProperty("city")]
            public string City { get; set; }

            [JsonProperty("region")]
            public string Region { get; set; }

            [JsonProperty("region_code")]
            public string RegionCode { get; set; }

            [JsonProperty("country")]
            public string Country { get; set; }

            [JsonProperty("country_code")]
            public string Country_code { get; set; }

            [JsonProperty("zip")]
            public string Zip { get; set; }

            [JsonProperty("phone_number")]
            public string PhoneNumber { get; set; }
        }



        /// ///////////////////////////////////////////////////////////////////////////////////////////////
        // FULFILLED


        //run once per/day
        // endpoint to use:  https://a.klaviyo.com/client/events/?company_id=HXtNd6
        public static Boolean Fulfilled_API(string FeedName, int RunElaspedMinutes = 30)
        {
            Boolean Returnvalue = true;
            string APIResult = "";
            DataTable dt = new DataTable();
            DataTable dtDetails;
            DataRow drDetails;
            KlaviyoOrder KO = new KlaviyoOrder();
            OrderItem IT; // = new OrderItem();
            string[] Cats;
            string[] Subcats;
            string[] Tastes;
            List<string> ItemNamesList;
            string OrderBrands;
            string APIName = "";
            string ItemCancel = "0";

            //Needs to be positive
            if (RunElaspedMinutes < 0)
            {
                RunElaspedMinutes = -RunElaspedMinutes;
            }

            if (FeedName.ToUpper() == "FULFILLED")
            {
                //dt = Helper.Sql_Misc_Fetch("SELECT orderno, confirmemail, total, store, ISNULL(DiscAmount,0) [DiscAmount], ISNULL(DiscType,'') [DiscType], ISNULL(DiscCode,'') [DiscCode], Replace(Convert(varchar(50),ShipDate ,102),'.','-') + 'T' + Convert(varchar(50),ShipDate ,108) [shipdate] FROM [hercust].[dbo].[orders] WHERE shipped = 1 AND SiteVersion=1 AND ShipDate >= Convert(date,dateadd(day, -1, GetDate())) Order By orderno");
                dt = Helper.Sql_Misc_Fetch("SELECT orderno, confirmemail, total, store, ISNULL(DiscAmount,0) [DiscAmount], ISNULL(DiscType,'') [DiscType], ISNULL(DiscCode,'') [DiscCode], Replace(Convert(varchar(50),ShipDate ,102),'.','-') + 'T12:00' [shipdate] FROM [hercust].[dbo].[orders] WHERE shipped = 1 AND SiteVersion=1 AND ShipDate >= Convert(date,dateadd(day, -1, GetDate())) Order By orderno");
                APIName = "Fulfilled Order";
            }
            else if (FeedName.ToUpper() == "PLACED")
            {
                dt = Helper.Sql_Misc_Fetch("SELECT magentoorderno [orderno], confirmemail, total, store, ISNULL(DiscAmount,0) [DiscAmount], ISNULL(DiscType,'') [DiscType], ISNULL(DiscCode,'') [DiscCode], Replace(Convert(varchar(50),ShipDate ,102),'.','-') + 'T' + Convert(varchar(50),ShipDate ,108) [shipdate] FROM [hercust].[dbo].[orders] "
                    + " WHERE retailerid = 0 AND siteversion = 1 AND TimeStamp >= dateadd(mi, -" + RunElaspedMinutes.ToString() + ", GetDate()) AND TimeStamp < GetDate() order by orderno ");
                APIName = "Placed Order";
            }
            else if (FeedName.ToUpper() == "CANCELLED")
            {
                dt = Helper.Sql_Misc_Fetch("SELECT orderno, confirmemail, total, store, ISNULL(DiscAmount,0) [DiscAmount], ISNULL(DiscType,'') [DiscType], ISNULL(DiscCode,'') [DiscCode], Replace(Convert(varchar(50),CancelledWhen ,102),'.','-') + 'T' + Convert(varchar(50),CancelledWhen,108) [shipdate] FROM [hercust].[dbo].[orders] "
                    + " WHERE cancel = 1 AND retailerid = 0 AND siteversion = 1 AND CancelledWhen >= dateadd(mi, -" + RunElaspedMinutes.ToString() + ", GetDate()) AND CancelledWhen < GetDate() order by orderno ");

                APIName = "Cancelled Order";
                ItemCancel = "1";
            }
            else if (FeedName.ToUpper() == "RETURNED")
            {
                dt = Helper.Sql_Misc_Fetch("SELECT DISTINCT  OIR.orderno, confirmemail, total, store, ISNULL(DiscAmount,0) [DiscAmount], ISNULL(DiscType, '')[DiscType], ISNULL(DiscCode, '') [DiscCode], Replace(Convert(varchar(50), OIR.ReturnDate, 102), '.', '-') + 'T' + Convert(varchar(50), OIR.ReturnDate, 108)[shipdate] "
                    + " FROM hercust..OrderItemReturns OIR "
                    + " INNER JOIN hercust..orders OO ON OO.OrderNo = OIR.OrderNo AND OO.Returned > 0 AND Siteversion = 1 "
                    + " WHERE OIR.ReturnDate BETWEEN dateadd(mi, -" + RunElaspedMinutes.ToString() + ", GetDate()) AND GetDate() ORDER BY 1 ");

                APIName = "Returned Order";
            }

            foreach (DataRow dr in dt.Rows )
            {
                dtDetails = new DataTable();        // clear it out
               
                dtDetails = Helper.Sql_Misc_Fetch("EXEC herroom..proc_get_klaviyo_info_byorderno_api @orderno = " + dr["orderno"].ToString() + ", @Cancel = " + ItemCancel);
                if (dtDetails.Rows.Count > 0)
                {
                   drDetails = dtDetails.Rows[0];
                    OrderBrands = "";

                    try
                    {

                        KO.data = new OrderData();
                        KO.data.Type = "event";
                        KO.data.OrderAttributes = new OrderAttributes();
                        KO.data.OrderAttributes.Properties = new OrderProperties();

                        KO.data.OrderAttributes.Metric = new OrderAttributeMetric();
                        KO.data.OrderAttributes.Metric.Data = new OrderAttributeMetricData();
                        KO.data.OrderAttributes.Metric.Data.Type = "metric";
                        KO.data.OrderAttributes.Metric.Data.Attributes = new OrderAttributeMetricDataAttributes();
                        KO.data.OrderAttributes.Metric.Data.Attributes.Name = APIName; // "Fulfilled Order";
                        KO.data.OrderAttributes.profile = new OrderProfile();
                        KO.data.OrderAttributes.profile.Data = new OrderProfileData();

                        KO.data.OrderAttributes.profile.Data.Attributes = new OrderProfileDataAttributes();
                        KO.data.OrderAttributes.profile.Data.Attributes.Properties = new OrderProfileDataAttributesProperties();
                        KO.data.OrderAttributes.profile.Data.Attributes.Email = dr["confirmemail"].ToString();
                        KO.data.OrderAttributes.profile.Data.Type = "profile";

                        KO.data.OrderAttributes.Properties.Orderno = dr["orderno"].ToString();

                        KO.data.OrderAttributes.Properties.Categories = new List<string>();
                        Cats = drDetails["OrderSubCategories"].ToString().Split(Char.Parse(","));
                        foreach (string Cat in Cats)
                        {
                            if (Cat.Trim().Length > 0)
                            {
                                KO.data.OrderAttributes.Properties.Categories.Add(Cat.Replace("&amp", "&"));
                            }
                        }

                        KO.data.OrderAttributes.Properties.Subcategories = new List<string>();
                        Subcats = drDetails["SubCategories"].ToString().Split(Char.Parse(","));
                        foreach (string Subcat in Subcats)
                        {
                            if (Subcat.Trim().Length > 0)
                            {
                                KO.data.OrderAttributes.Properties.Subcategories.Add(Subcat.Replace("&amp;", "&").Replace("'", ""));
                            }
                        }

                        KO.data.OrderAttributes.Properties.Tastes = new List<string>();
                        Tastes = drDetails["ordertastes"].ToString().Split(Char.Parse(","));
                        foreach (string Taste in Tastes)
                        {
                            if (Taste.Trim().Length > 0)
                            {
                                KO.data.OrderAttributes.Properties.Tastes.Add(Taste);
                            }
                        }

                        KO.data.OrderAttributes.Properties.Brands = drDetails["ManufacturerName"].ToString();
                        KO.data.OrderAttributes.Properties.Onsale = drDetails["orderonsale"].ToString();
                        KO.data.OrderAttributes.Properties.Braband = drDetails["orderbraband"].ToString();
                        KO.data.OrderAttributes.Properties.Braddplus = drDetails["orderddplus"].ToString();

                        KO.data.OrderAttributes.Properties.DiscountCode = dr["disccode"].ToString();
                        KO.data.OrderAttributes.Properties.DiscountType = dr["disctype"].ToString();
                        KO.data.OrderAttributes.Properties.DiscountValue = decimal.Parse(dr["discamount"].ToString());

                        KO.data.OrderAttributes.Value = decimal.Parse(dr["total"].ToString());
                        KO.data.OrderAttributes.Value_currency = "USD";

                        //Items - one in json, not for each item
                        KO.data.OrderAttributes.Properties.Items = new List<OrderItem>();
                        string[] ItemCats;
                        //string[] ItemSubCats;
                        ItemNamesList = new List<string>();

                        foreach (DataRow drDetail in dtDetails.Rows)
                        {
                            //Items
                            IT = new OrderItem();
                            IT.Categories = new List<string>();
                            IT.ProductName = drDetail["ProductName"].ToString();
                            IT.SKU = drDetail["upc"].ToString();
                            IT.ProductID = drDetails["stylenumber"].ToString();
                            IT.Size = drDetail["size"].ToString();
                            IT.Color = drDetail["colorname"].ToString();
                            IT.Quantity = int.Parse(drDetail["qty"].ToString());
                            IT.ItemPrice = decimal.Parse(drDetail["unitprice"].ToString());
                            IT.onsale = drDetail["onsale"].ToString();
                            IT.braband = drDetail["orderbraband"].ToString();
                            IT.braddplus = drDetail["braddplus"].ToString();
                            IT.Brands = drDetail["manufacturername"].ToString();
                            IT.ProductURL = drDetail["styleurl"].ToString();
                            IT.ImageURL = drDetail["imageurl"].ToString();


                            ItemCats = drDetail["subcategories"].ToString().Split(Char.Parse(","));
                            foreach (string Cat in ItemCats)
                            {
                                if (Cat.Trim().Length > 0)
                                {
                                    IT.Categories.Add(Cat.Replace("&amp", "&"));
                                }
                            }

                            KO.data.OrderAttributes.Properties.Items.Add(IT);

                            ItemNamesList.Add(drDetail["ProductName"].ToString());

                            OrderBrands += drDetail["manufacturername"].ToString() + ",";
                        }

                        if (OrderBrands.Length > 2)
                        {
                            OrderBrands.TrimEnd(Char.Parse(","));
                        }
                        KO.data.OrderAttributes.Properties.Brands = OrderBrands;
                        KO.data.OrderAttributes.Properties.ItemNames = ItemNamesList;
                        KO.data.OrderAttributes.Time = dr["shipdate"].ToString();

                        var TM_Json = JsonConvert.SerializeObject(KO, Formatting.Indented, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });

                        Console.WriteLine(TM_Json);

                        APIResult = Helper.MagentoApiPush_Klaviyo("event", TM_Json, dr["store"].ToString());
                        if (MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine(dr["confirmemail"].ToString() + " :: " + APIResult);
                            Console.WriteLine(" ------------------ ");
                        }
                    }
                    catch(Exception ex)
                    {
                        if (MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine("Klaviyo API ERROR: " + ex.ToString());
                        }
                        Returnvalue = false;
                    }
                }

            };


         
            return Returnvalue;
        }


        /// ///////////////////////////////////////////////////////////////////////////////////////////////
        // Placedmin_AP

       
       

    }
}
