using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace Magento_MCP.MagentoModels.ExtensionAttributeModels
{
	public partial class InvoiceItemExtensionAttribute
	{
		[JsonProperty("vertex_tax_codes")]
		public List<string> VertexTaxCodes { get; set; }

		[JsonProperty("invoice_text_codes")]
		public List<string> InvoiceTextCodes { get; set; }

		[JsonProperty("tax_codes")]
		public List<string> TaxCodes { get; set; }
	}
}
