using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Magento_MCP.MagentoModels
{
	class MediaGalleryEntries
	{
		public int id { get; set; }
		public string media_type { get; set; }
		public string label { get; set; }
		public int position { get; set; }
		public bool disabled { get; set; }
		public List<string> types { get; set; } //"types": ["image","small_image","thumbnail","swatch_image"],
		public string file { get; set; }
		public Content content { get; set; }
		public ExtensionAttributeModels.MediaGalleryExtAttr extension_attributes { get; set; }
	}
}
