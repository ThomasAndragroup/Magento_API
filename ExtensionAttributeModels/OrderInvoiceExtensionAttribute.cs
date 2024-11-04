using Newtonsoft.Json;
using Magento_MCP.MagentoModels.HelperModels;

namespace Magento_MCP.MagentoModels.ExtensionAttributeModels
{
	public partial class OrderInvoiceExtensionAttribute
	{
		[JsonProperty("base_customer_balance_amount")]
		public float BaseCustomerBalanceAmount { get; set; }

		[JsonProperty("customer_balance_amount")]
		public float CustomerBalanceAmount { get; set; }

		[JsonProperty("base_gift_cards_amount")]
		public float BaseGiftCardsAmount { get; set; }

		[JsonProperty("gift_cards_amount")]
		public float GiftCardsAmount { get; set; }

		[JsonProperty("gw_base_price")]
		public string GwBasePrice { get; set; }

		[JsonProperty("gw_price")]
		public string GwPrice { get; set; }

		[JsonProperty("gw_items_base_price")]
		public string GwItemsBasePrice { get; set; }

		[JsonProperty("gw_items_price")]
		public string GwItemsPrice { get; set; }

		[JsonProperty("gw_card_base_price")]
		public string GwCardBasePrice { get; set; }

		[JsonProperty("gw_card_price")]
		public string GwCardPrice { get; set; }

		[JsonProperty("gw_base_tax_amount")]
		public string GwBaseTaxAmount { get; set; }

		[JsonProperty("gw_tax_amount")]
		public string GwTaxAmount { get; set; }

		[JsonProperty("gw_items_base_tax_amount")]
		public string GwItemsBaseTaxAmount { get; set; }

		[JsonProperty("gw_items_tax_amount")]
		public string GwItemsTaxAmount { get; set; }

		[JsonProperty("gw_card_base_tax_amount")]
		public string GwCardBaseTaxAmount { get; set; }

		[JsonProperty("gw_card_tax_amount")]
		public string GwCardTaxAmount { get; set; }

		//[JsonProperty("vertex_tax_calculation_shipping_address")]
		//public OrderInvoiceVertexShippingAddress VertexTaxCalculationShippingAddress { get; set; }

		//[JsonProperty("vertex_tax_calculation_billing_address")]
		//public OrderInvoiceVertexShippingAddress VertexTaxCalculationBillingAddress { get; set; }

		//[JsonProperty("vertex_tax_calculation_order")]
		//public Magento_MCP.MagentoModels.OrdersModels.Entity VertexTaxCalculationOrder { get; set; }
	}
}
