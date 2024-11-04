using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace Magento_MCP.MagentoModels.HelperModels
{
	class OrderPostHelpers
	{
	}
	public class PostBase
	{
		[JsonProperty("notify")]
		public bool Notify { get; set; }

		[JsonProperty("appendComment")]
		public bool AppendComment { get; set; }
	}
	public class PostItemBase
	{
		[JsonProperty("order_item_id")]
		public long OrderItemId { get; set; }

		[JsonProperty("qty")]
		public long Qty { get; set; }
	}


}
