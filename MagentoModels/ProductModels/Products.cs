using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Magento_MCP.MagentoModels.ExtensionAttributeModels;
using Newtonsoft.Json;

namespace Magento_MCP.MagentoModels.ProductModels
{
	class Products
	{
		public Product product { get; set; }
	}

	class Product
	{
		public string sku { get; set; }
		public string name { get; set; }
		public int attribute_set_id { get; set; }
		public ProductStatus status { get; set; }
		public ProductVisibility visibility { get; set; }
		public string type_id { get; set; }
		public double price { get; set; }
		public double weight { get; set; }
		public ProductPostExtAttr extension_attributes { get; set; }
		public List<ProductLinks> productLinks { get; set; }
		public List<ProductOption> options { get; set; }
		public List<MediaGalleryEntries> media_gallery_entries { get; set; }
		public List<TierPrice> tier_prices { get; set; }
		public List<CustomAttribute> custom_attributes { get; set; }   
        //public Boolean saveOptions { get; set; }
	}

	class ProductResponse	:	Product
	{
		public int id { get; set; }
		public string created_at { get; set; }
		public string updated_at { get; set; }
	}
}
