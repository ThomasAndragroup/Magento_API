using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Magento_MCP.MagentoModels.ExtensionAttributeModels;

namespace Magento_MCP.MagentoModels.HelperModels
{
	public class RmaComment
	{
		[JsonProperty("comment")]
		public string CommentComment { get; set; }

		[JsonProperty("rma_entity_id")]
		public long RmaEntityId { get; set; }

		[JsonProperty("created_at")]
		public string CreatedAt { get; set; }

		[JsonProperty("entity_id")]
		public long EntityId { get; set; }

		[JsonProperty("customer_notified")]
		public bool CustomerNotified { get; set; }

		[JsonProperty("visible_on_front")]
		public bool VisibleOnFront { get; set; }

		[JsonProperty("status")]
		public string Status { get; set; }

		[JsonProperty("admin")]
		public bool Admin { get; set; }

		[JsonProperty("extension_attributes")]
		public BlankExtensionAttribute ExtensionAttributes { get; set; }

		[JsonProperty("custom_attributes")]
		public List<CustomAttribute> CustomAttributes { get; set; }
	}
}
