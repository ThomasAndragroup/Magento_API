using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Magento_MCP.MagentoModels.ExtensionAttributeModels;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace Magento_MCP.MagentoModels.HelperModels
{
	public class ShippingItem 
	{
		[JsonProperty("additional_data")]
		public string AdditionalData { get; set; }

		[JsonProperty("description")]
		public string Description { get; set; }

		[JsonProperty("entity_id")]
		public long EntityId { get; set; }

		[JsonProperty("name")]
		public string Name { get; set; }

		[JsonProperty("parent_id")]
		public long ParentId { get; set; }

		[JsonProperty("price")]
		public decimal Price { get; set; }

		[JsonProperty("product_id")]
		public long ProductId { get; set; }

		[JsonProperty("row_total")]
		public decimal RowTotal { get; set; }

		[JsonProperty("sku")]
		public string Sku { get; set; }

		[JsonProperty("weight")]
		public decimal Weight { get; set; }

		/// <summary>
		/// By request of creatuity this has been set to a blank object
		/// </summary>
		[JsonProperty("extension_attributes")]
		public BlankExtensionAttribute ExtensionAttributes { get; set; }

		[JsonProperty("order_item_id")]
		public long OrderItemId { get; set; }

		[JsonProperty("qty")]
		public long Qty { get; set; }
	}
}
