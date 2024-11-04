using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Magento_MCP.MagentoModels.ExtensionAttributeModels;

namespace Magento_MCP.MagentoModels.HelperModels
{
public class RmaItem
	{
		[JsonProperty("entity_id")]
		public long EntityId { get; set; }

		[JsonProperty("rma_entity_id")]
		public long RmaEntityId { get; set; }

		[JsonProperty("order_item_id")]
		public long OrderItemId { get; set; }

		[JsonProperty("qty_requested")]
		public long QtyRequested { get; set; }

		[JsonProperty("qty_authorized")]
		public long QtyAuthorized { get; set; }

		[JsonProperty("qty_approved")]
		public long QtyApproved { get; set; }

		[JsonProperty("qty_returned")]
		public long QtyReturned { get; set; }

		[JsonProperty("reason")]
		public string Reason { get; set; }

		[JsonProperty("condition")]
		public string Condition { get; set; }

		[JsonProperty("resolution")]
		public string Resolution { get; set; }

		[JsonProperty("status")]
		public string Status { get; set; }

        [JsonProperty("htorderno")]
        public string htorder { get; set; }

        [JsonProperty("extension_attributes")]
		public BlankExtensionAttribute ExtensionAttributes { get; set; }
	}
}
