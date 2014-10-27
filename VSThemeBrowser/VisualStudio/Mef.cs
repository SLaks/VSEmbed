using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.ComponentModel.Composition.Primitives;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.Text.Storage;
using Microsoft.VisualStudio.Utilities;
using VSThemeBrowser.Controls;

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

			// SLaks: Needed for VsUndoHistoryRegistry (which doesn't actually work), VsWpfKeyboardTrackingService, & probably others
			"Microsoft.VisualStudio.Editor.Implementation",

			// SLaks: Needed for IVsHierarchyItemManager, used by peek providers
			"Microsoft.VisualStudio.Shell.TreeNavigation.HierarchyProvider"
		};

		// I need to specify a full name to load from the GAC.
		// The version is added by my AssemblyResolve handler.
		const string FullNameSuffix = ", Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL";
		static IEnumerable<ComposablePartCatalog> GetCatalogs() {
			return EditorComponents.Select(c => new AssemblyCatalog(Assembly.Load(c + FullNameSuffix)))
					.Concat(GetRoslynCatalogs())
					.Concat(new[] { new AssemblyCatalog(typeof(Mef).Assembly) });
		}
		static IEnumerable<ComposablePartCatalog> GetRoslynCatalogs() {
			if (VsLoader.RoslynAssemblyPath == null)
				return new ComposablePartCatalog[0];
			return Directory.EnumerateFiles(VsLoader.RoslynAssemblyPath, "Microsoft.CodeAnalysis*.dll")	// Leave out the . to catch Microsoft.CodeAnalysis.dll too
				.Select(p => new AssemblyCatalog(Assembly.LoadFile(p)))
				.Concat(new[] { new AssemblyCatalog(Assembly.Load("Microsoft.VisualStudio.LanguageServices")) });
		}

		public static readonly CompositionContainer Container =
			new CompositionContainer(new AggregateCatalog(GetCatalogs()));
		static Mef() {
			// Copied from Microsoft.VisualStudio.ComponentModelHost.ComponentModel.DefaultCompositionContainer
			Container.ComposeExportedValue<SVsServiceProvider>(
				new VsServiceProviderWrapper(ServiceProvider.GlobalProvider));

			// Needed because VsUndoHistoryRegistry tries to create IOleUndoManager from ILocalRegistry, which I presumably cannot do.
			Container.ComposeExportedValue((ITextUndoHistoryRegistry)EditorUtils.EditorUtilsFactory.CreateBasicUndoHistoryRegistry());
		}

		// Microsoft.VisualStudio.Editor.Implementation.DataStorage uses COM services
		// that read the user's color settings, which I cannot easily duplicate.  The
		// editor reads MEF-exported defaults in EditorFormatMap, so I do not need to
		// implement this at all unless I want to allow user customization.
		sealed class SimpleDataStorage : IDataStorage {
			public bool TryGetItemValue(string itemKey, out ResourceDictionary itemValue) {
				itemValue = null;
				return false;
			}
		}
		[Export(typeof(IDataStorageService))]
		sealed class DataStorageService : IDataStorageService {
			readonly IDataStorage instance = new SimpleDataStorage();
			public IDataStorage GetDataStorage(string storageKey) { return instance; }
		}

		[Export(typeof(IKeyProcessorProvider))]
		[TextViewRole(PredefinedTextViewRoles.Interactive)]
		[ContentType("text")]
		[Name("Simple KeyProcessor")]
		sealed class SimpleKeyProcessorProvider : IKeyProcessorProvider {
			[Import]
			public IEditorOperationsFactoryService EditorOperationsFactory { get; set; }
			[Import]
			public ITextUndoHistoryRegistry UndoHistoryRegistry { get; set; }
			public KeyProcessor GetAssociatedProcessor(IWpfTextView wpfTextView) {
				return new SimpleKeyProcessor(wpfTextView, EditorOperationsFactory.GetEditorOperations(wpfTextView), UndoHistoryRegistry.GetHistory(wpfTextView.TextBuffer));
			}
		}
	}
}
