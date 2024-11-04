using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Magento_MCP.MagentoModels.ProductModels
{
	class ProductLinkOptions
	{
		/// <summary>
		/// Use this as parent so we have a parent node called options
		/// url to post a link will be V1/configurable-products/{parentsku}/options
		/// </summary>
		public ProductLinkOption option { get; set; }
	}
	class ProductLinkOption
	{
		/// <summary>
		/// for example attribute id: 208
		/// </summary>
		public string attribute_id { get; set; }
		/// <summary>
		/// label for example is size or an attribute
		/// </summary>
		public string label { get; set; }
		public int position { get; set; }
		public bool is_use_default { get; set; }
		/// <summary>
		/// Value index should be set to a unique number for example a current timestamp
		/// </summary>
		public List<ProductLinkValue> values { get; set; }
	}
}
