using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Magento_MCP.MagentoModels.ExtensionAttributeModels;

namespace Magento_MCP.MagentoModels.HelperModels
{
	public class CommentBase
	{
		[JsonProperty("comment")]
		public string Comment { get; set; }

		[JsonProperty("is_visible_on_front")]
		public long IsVisibleOnFront { get; set; }

		/// <summary>
		/// todo: need to update this extension attribute
		/// object (sales-data-shipment-comment-extension-interface)
		/// ExtensionInterface class for @see \Magento\Sales\Api\Data\ShipmentCommentInterface
		/// </summary>
		[JsonProperty("extension_attributes")]
		public BlankExtensionAttribute ExtensionAttributes { get; set; }
	}
	public class Comment : CommentBase
	{
		[JsonProperty("is_customer_notified")]
		public long IsCustomerNotified { get; set; }

		[JsonProperty("parent_id")]
		public long ParentId { get; set; }

		[JsonProperty("created_at")]
		public string CreatedAt { get; set; }

		[JsonProperty("entity_id")]
		public long EntityId { get; set; }
	}

	public class InvoiceComment : Comment
	{
		[JsonProperty("entity_name", NullValueHandling = NullValueHandling.Ignore)]
		public string EntityName { get; set; }

		[JsonProperty("status", NullValueHandling = NullValueHandling.Ignore)]
		public string Status { get; set; }
	}
	public class OrderInvoicePostComment : CommentBase
	{

	}
}
