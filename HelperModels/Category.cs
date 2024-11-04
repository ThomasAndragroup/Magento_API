using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Magento_MCP.MagentoModels.ExtensionAttributeModels;

namespace Magento_MCP.MagentoModels.HelperModels
{
	public class Category
	{
		[JsonProperty("id")]
		public long Id { get; set; }

		[JsonProperty("is_active")]
		public String IsActive { get; set; }

		[JsonProperty("include_in_menu")]
		public string include_in_menu { get; set; }

		[JsonProperty("parent_id")]
		public long ParentId { get; set; }

		[JsonProperty("name")]
		public string Name { get; set; }

		[JsonProperty("customAttributes")]
		public CustomAttribute[] CustomAttributes { get; set; }
	}
}
 