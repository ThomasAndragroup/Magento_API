using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Magento_MCP.MagentoModels.ExtensionAttributeModels;
using Newtonsoft.Json;

namespace Magento_MCP.MagentoModels.HelperModels
{
	public class TrackBase
	{
		[JsonProperty("track_number")]
		public string TrackNumber { get; set; }

		[JsonProperty("title")]
		public string Title { get; set; }

		[JsonProperty("carrier_code")]
		public string CarrierCode { get; set; }
	}
	public class Track : TrackBase
	{
		[JsonProperty("order_id")]
		public long OrderId { get; set; }

		[JsonProperty("created_at")]
		public string CreatedAt { get; set; }

		[JsonProperty("entity_id")]
		public long EntityId { get; set; }

		[JsonProperty("parent_id")]
		public long ParentId { get; set; }

		[JsonProperty("updated_at")]
		public string UpdatedAt { get; set; }

		[JsonProperty("weight")]
		public decimal Weight { get; set; }

		[JsonProperty("qty")]
		public long Qty { get; set; }

		[JsonProperty("description")]
		public string Description { get; set; }

		/// <summary>
		/// left blank object on puropose on request to creatuity
		/// </summary>
		[JsonProperty("extension_attributes")]
		public BlankExtensionAttribute ExtensionAttributes { get; set; }
	}

    public class RmaTrackIncoming
    {
        [JsonProperty("rma_track_number")]
        public RmaTrack RmaTrackNumber { get; set; }
    }

	public class RmaTrack
	{
		[JsonProperty("entity_id")]
		public long EntityId { get; set; }

		[JsonProperty("rma_entity_id")]
		public long RmaEntityId { get; set; }

		[JsonProperty("track_number")]
		public string TrackNumber { get; set; }

		[JsonProperty("carrier_title")]
		public string CarrierTitle { get; set; }

		[JsonProperty("carrier_code")]
		public string CarrierCode { get; set; }

		[JsonProperty("extension_attributes")]
		public BlankExtensionAttribute ExtensionAttributes { get; set; }
	}

	public class ShippingOrderPostTrack : TrackBase
	{
		[JsonProperty("extension_attributes")]
		public BlankExtensionAttribute ExtensionAttributes { get; set; }
	}
}
