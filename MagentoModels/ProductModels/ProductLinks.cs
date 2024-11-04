using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Magento_MCP.MagentoModels.ProductModels
{
	class ProductLinks
	{
		public string sku { get; set; }
		public string link_type { get; set; }
		public string linked_product_sku { get; set; }
		public string linked_product_type { get; set; }
		public int position { get; set; }
		public ExtensionAttributeModels.ProductPostExtAttr extension_attributes { get; set; }
	}
}
