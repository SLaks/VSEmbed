using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using VSThemeBrowser.VisualStudio;

namespace VSThemeBrowser.Controls {
	public class ConditionalResourceDictionary : ResourceDictionary {
		public ConditionalResourceDictionary() {
			var observableChildren = (ObservableCollection<ResourceDictionary>)MergedDictionaries;
			observableChildren.CollectionChanged += MergedDictionaries_CollectionChanged;
		}

		private void MergedDictionaries_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e) {
			if (e.NewItems == null)
				return;
			foreach (var child in e.NewItems.OfType<ConditionalResourceDictionary>()) {
				if (child.IsRelevant)
					child.Clear();
			}
		}

		///<summary>The minimum (inclusive) version of Visual Studio to load the ResourceDictionary for.</summary>
		public int? MinVersion { get; set; }
		///<summary>The maximum (inclusive) version of Visual Studio to load the ResourceDictionary for.</summary>
		public int? MaxVersion { get; set; }

		private bool IsRelevant {
			get {
				return !(MinVersion > VsLoader.VsVersion.Major || MaxVersion < VsLoader.VsVersion.Major);
			}
		}

		protected override void OnGettingValue(object key, ref object value, out bool canCache) {
			if (IsRelevant)
				base.OnGettingValue(key, ref value, out canCache);
			else {
				canCache = true;
				value = null;
			}
		}
	}
}
