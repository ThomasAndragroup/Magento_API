using Newtonsoft.Json;


namespace Magento_MCP.MagentoModels.ExtensionAttributeModels
{
	public class OrderInvoiceVertexShippingAddressExtensionAttribute
	{
		[JsonProperty("vertex_vat_country_code")]
		public string VertexVatCountryCode { get; set; }
	}
}
