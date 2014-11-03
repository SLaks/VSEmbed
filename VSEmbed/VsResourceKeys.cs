using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VSEmbed {
	///<summary>Contains resource keys for resources that change based on Visual Studio versions.</summary>
	///<remarks>
	/// Dev14 has standard styles for basic controls, keyed as VsComboBoxStyleKey etc.
	/// For earlier VS versions, I have my own themed styles for these controls.  This
	/// class contains resource keys that will return the VS style or my custom styles
	/// depending on the loaded VS version.
	///</remarks>
	public static class VsResourceKeys {
		public static bool HasDefaultStyles { get { return VsLoader.VsVersion.Major >= 14; } }

		public static string ComboBoxStyleKey {
			get { return HasDefaultStyles ? "VsComboBoxStyleKey" : "SLaks.ComboBoxStyleKey"; }
		}
		public static string ButtonStyleKey {
			get { return HasDefaultStyles ? "VsButtonStyleKey" : "SLaks.ButtonStyleKey"; }
		}
		public static string CheckBoxStyleKey {
			get { return HasDefaultStyles ? "VsCheckBoxStyleKey" : "SLaks.CheckBoxStyleKey"; }
		}
	}
}
