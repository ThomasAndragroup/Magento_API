using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Magento_MCP.MagentoModels.ValueModels
{
	class ProductOptionValue
	{
		public string title { get; set; }
		public int sort_order { get; set; }
		public int price { get; set; }
		public string price_type { get; set; }
		public string sku { get; set; }
		public int option_type_id { get; set; }
	}
}
