using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Magento_MCP.MagentoModels.ExtensionAttributeModels
{
	public abstract class TierPriceExtAttr
	{
		public int percentage_value { get; set; }
		public int website_id { get; set; }
	}
}
