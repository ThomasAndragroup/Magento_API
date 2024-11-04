using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace Magento_MCP.MagentoModels.ExtensionAttributeModels
{
	public class RefundArgumentsExtAttr
	{
		[JsonProperty("return_to_stock_items")]
		public List<long> ReturnToStockItems { get; set; }

		[JsonProperty("is_store_credit")]
		public bool IsStoreCredit { get; set; }
	}
}
