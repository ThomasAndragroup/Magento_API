using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Globalization;
using Magento_MCP.MagentoModels.HelperModels;
using Magento_MCP.MagentoModels.ExtensionAttributeModels;
using Newtonsoft.Json.Converters;


// DO NOT NEED THIS, SALESMODELS IS WORKING FOR  THIS

namespace Magento_MCP.MagentoModels.SalesModels
{
    public partial class Rma
    {
        [JsonProperty("rmaDataObject")]
        public RmaDataObject RmaDataObject { get; set; }
    }

    public partial class RmaIncoming
    {
        [JsonProperty("rma")]
        public RmaDataObject RmaDataObject { get; set; }
    }

    public partial class RmaDataObject
    {
        [JsonProperty("increment_id")]
        public string IncrementId { get; set; }

        [JsonProperty("entity_id")]
        public long EntityId { get; set; }

        [JsonProperty("order_id")]
        public long OrderId { get; set; }

        [JsonProperty("order_increment_id")]
        public string OrderIncrementId { get; set; }

        [JsonProperty("store_id")]
        public long StoreId { get; set; }

        [JsonProperty("customer_id")]
        public long CustomerId { get; set; }

        [JsonProperty("date_requested")]
        public string DateRequested { get; set; }

        [JsonProperty("customer_custom_email")]
        public string CustomerCustomEmail { get; set; }

        [JsonProperty("items")]
        public List<RmaItem> Items { get; set; }

        [JsonProperty("status")]
        public string Status { get; set; }

        [JsonProperty("comments")]
        public List<RmaComment> Comments { get; set; }

        [JsonProperty("tracks")]
        public List<RmaTrack> Tracks { get; set; }

        [JsonProperty("extension_attributes")]
        public BlankExtensionAttribute ExtensionAttributes { get; set; }

        [JsonProperty("custom_attributes")]
        public List<CustomAttribute> CustomAttributes { get; set; }
    }


	public partial class Rma
	{
		//public static Rma FromJson(string json) => JsonConvert.DeserializeObject<Rma>(json, SalesModels.Converter.Settings);
	}

	public static class Serialize
	{
		//public static string ToJson(this Rma self) => JsonConvert.SerializeObject(self, SalesModels.Converter.Settings);
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
