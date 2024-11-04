using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Magento_MCP.MagentoModels
{
	class Options
	{
		public Option option { get; set; }
	}
	class Option
	{
		public string label { get; set; }
		public string value { get; set; }
		public int sort_order { get; set; }
		public bool is_default { get; set; }
		public List<StoreLabels> store_labels { get; set; }
	}
	class ProductOption
	{
		public string product_sku { get; set; }
		public int option_id { get; set; }
		public string title { get; set; }
		public string type { get; set; }
		public int sort_order { get; set; }
		public bool is_required { get; set; }
		public double price { get; set; }
		public string price_type { get; set; }
		public string sku { get; set; }
		public string file_extension { get; set; }
		public int max_characters { get; set; }
		public int image_size_x { get; set; }
		public int image_size_y { get; set; }
		public List<ValueModels.ProductOptionValue> values { get; set; }
		public ExtensionAttributeModels.ProductOptionExtAttr extension_attributes { get; set; }

	}
}
