using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Storage;

namespace VSThemeBrowser.VisualStudio.Exports {
	// This file contains trivially simply MEF exports

	[Export(typeof(IExtensionErrorHandler))]
	sealed class SimpleErrorReporter : IExtensionErrorHandler {
		public void HandleError(object sender, Exception exception) {
			MessageBox.Show(exception.ToString(), "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
		}
	}

	// Microsoft.VisualStudio.Editor.Implementation.DataStorage uses COM services
	// that read the user's color settings, which I cannot easily duplicate.  The
	// editor reads MEF-exported defaults in EditorFormatMap, so I do not need to
	// implement this at all unless I want to allow user customization.
	sealed class SimpleDataStorage : IDataStorage {
		readonly ResourceDictionary plainTextValue;
		public SimpleDataStorage(ResourceDictionary plainTextValue) {
			this.plainTextValue = plainTextValue;
		}

		public bool TryGetItemValue(string itemKey, out ResourceDictionary itemValue) {
			switch (itemKey) {
				case "Plain Text":
					itemValue = plainTextValue;
					return plainTextValue != null;
				case "Collapsible Text (Collapsed)":
					// This is used by CollapsedAdornmentProvider, and has no default value.
					itemValue = new ResourceDictionary { { "Foreground", Brushes.Black } };
					return true;
				default:
					itemValue = null;
					return false;
			}
		}
	}

	[Export(typeof(IDataStorageService))]
	sealed class DataStorageService : IDataStorageService {
		readonly IDataStorage defaultInstance = new SimpleDataStorage(null);
		readonly IDataStorage messageInstance = new SimpleDataStorage(new ResourceDictionary {
			{ "Typeface", new Typeface(SystemFonts.MessageFontFamily, SystemFonts.MessageFontStyle, SystemFonts.MessageFontWeight, FontStretches.Normal) },
			{ "FontRenderingSize", SystemFonts.MessageFontSize }
		});

		public IDataStorage GetDataStorage(string storageKey) {
			switch (storageKey) {
				// Don't use a fixed-width font for tooltips & IntelliSense
				case "tooltip":
				case "completion":
					return messageInstance;
				default:
					return defaultInstance;
			}
		}
	}
}
