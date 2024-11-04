using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Magento_MCP.MagentoModels.ExtensionAttributeModels
{
	public class SceneSevenImages 
	{
		public string primary_image { get; set; }
		public string back_side_view_image { get; set; }
		public string front_side_view_image { get; set; }
		public string measuring_image { get; set; }
		public List<string> see_it_under_images { get; set; }
		public string temperature_image { get; set; }
		public List<string> gallery_images { get; set; }
		public string swatch_image { get; set; }
		public List<string> optional_images { get; set; }
		public List<string> spin_images { get; set; }

        public string tombstone_front_image { get; set; }

        public string tombstone_back_image { get; set; }

    }
}
