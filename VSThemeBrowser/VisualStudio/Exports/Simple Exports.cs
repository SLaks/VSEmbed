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
		public bool TryGetItemValue(string itemKey, out ResourceDictionary itemValue) {
			switch (itemKey) {
				// This is used by CollapsedAdornmentProvider, and has no default value.
				// However, MEF doesn't use my DataStorageService, so this doesn't work.
				case "Collapsible Text (Collapsed)":
					itemValue = new ResourceDictionary { { "Foreground", Brushes.DarkSlateBlue } };
					return true;
				default:
					itemValue = null;
					return false;
			}
		}
	}
	[Export(typeof(IDataStorageService))]
	sealed class DataStorageService : IDataStorageService {
		readonly IDataStorage instance = new SimpleDataStorage();
		public IDataStorage GetDataStorage(string storageKey) { return instance; }
	}

}
