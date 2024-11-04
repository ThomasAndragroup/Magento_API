using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Magento_MCP.MagentoModels.ExtensionAttributeModels;

namespace Magento_MCP.MagentoModels.HelperModels
{
	public class Package
	{
		/// <summary>
		/// todo: need to update this extension attribute 
		/// https://magento.redoc.ly/2.4.3-admin/tag/shipment#operation/salesShipmentRepositoryV1SavePost
		/// </summary>
		[JsonProperty("extension_attributes")]
		public BlankExtensionAttribute ExtensionAttributes { get; set; }
	}
}
