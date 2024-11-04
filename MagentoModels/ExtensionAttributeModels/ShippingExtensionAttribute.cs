using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace Magento_MCP.MagentoModels.ExtensionAttributeModels
{
	public class ShippingExtensionAttribute
	{
		[JsonProperty("source_code")]
		public string SourceCode { get; set; }
	}
}
