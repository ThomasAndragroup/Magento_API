using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Magento_MCP.MagentoModels.ExtensionAttributeModels
{
	class ProductPostExtAttr
	{
		public List<int> website_ids { get; set; }
		//public List<CategoryLinks> category_links { get; set; }
		//public StockData stock_item { get; set; }
		public SceneSevenImages scene_seven_images { get; set; }
	}
}
