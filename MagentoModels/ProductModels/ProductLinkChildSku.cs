using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Magento_MCP.MagentoModels.ProductModels
{
	class ProductLinkChildSku
	{
		/// <summary>
		/// url to POST will be V1/configurable-products/{parentsku}/child
		/// one child per request
		/// </summary>
		public string childSku { get; set; }
	}
}
