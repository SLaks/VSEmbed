using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.ComponentModel.Composition.Primitives;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;

namespace VSThemeBrowser.VisualStudio {
	///<summary>Creates the MEF composition container used by the editor services.</summary>
	/// <remarks>Stolen, with much love and gratitude, from @JaredPar's EditorUtils.</remarks>
	public static class Mef {
		private static readonly string[] EditorComponents = {
			// Core editor components
			"Microsoft.VisualStudio.Platform.VSEditor",

			// Not entirely sure why this is suddenly needed
			"Microsoft.VisualStudio.Text.Internal",

			// Must include this because several editor options are actually stored as exported information 
			// on this DLL.  Including most importantly, the tabsize information
			"Microsoft.VisualStudio.Text.Logic",

			// Include this DLL to get several more EditorOptions including WordWrapStyle
			"Microsoft.VisualStudio.Text.UI",

			// Include this DLL to get more EditorOptions values and the core editor
			"Microsoft.VisualStudio.Text.UI.Wpf",

			// SLaks: Needed for VsUndoHistoryRegistry, VsWpfKeyboardTrackingService, & probably others
			"Microsoft.VisualStudio.Editor.Implementation"
		};

		// I need to specify a full name to load from the GAC.
		// The version is added by my AssemblyResolve handler.
		const string FullNameSuffix = ", Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL";
        static IEnumerable<ComposablePartCatalog> GetCatalogs() {
			return EditorComponents.Select(c => new AssemblyCatalog(Assembly.Load(c + FullNameSuffix)));
		}
		public static readonly CompositionContainer Container = 
			new CompositionContainer(new AggregateCatalog(GetCatalogs()));
		static Mef() {
			Container.ComposeExportedValue<SVsServiceProvider>(
				new VsServiceProviderWrapper(ServiceProvider.GlobalProvider));
		}
	}
}
