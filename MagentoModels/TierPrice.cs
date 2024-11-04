using System;
using Newtonsoft.Json;
using Magento_MCP.MagentoCore;
using Magento_MCP.MagentoJsonWriter;

namespace Magento_MCP.MagentoModels
{
	public abstract class TierPrice
	{
		/// <summary>
		/// customer group id
		/// </summary>
		/// <remarks></remarks>
		public int customer_group_id { get; set; }
		public int qty { get; set; }
		public int value { get; set; }
		public ExtensionAttributeModels.TierPriceExtAttr extension_attributes { get; set; }

	}
}
