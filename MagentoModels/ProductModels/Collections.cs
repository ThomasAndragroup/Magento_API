using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Magento_MCP.MagentoModels.ProductModels
{
	class Collections
	{
		public Collection collection { get; set; }
	}
	class Collection
	{
		public string code { get; set; }
		public string name { get; set; }
		public string logo { get; set; }
		public int brand_id { get; set; }
		public List<int> store_id { get; set; }
		public string title_tag { get; set; }
		public string banner_section { get; set; }
	}
	class CollectionResponse : Collection
	{
		public int id { get; set; }
	}
}
