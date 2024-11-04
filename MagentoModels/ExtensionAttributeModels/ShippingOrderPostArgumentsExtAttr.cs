
using Newtonsoft.Json;

namespace Magento_MCP.MagentoModels.ExtensionAttributeModels
{
	public class ShippingOrderPostArgumentsExtAttr
	{
		[JsonProperty("source_code")]
		public string SourceCode { get; set; }
	}
}
