using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.ComponentModel.Composition.Primitives;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Media;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.Text.Storage;
using Microsoft.VisualStudio.Utilities;
using VSThemeBrowser.Controls;

namespace VSThemeBrowser.VisualStudio {
	///<summary>Creates the MEF composition container used by the editor services.</summary>
	/// <remarks>Stolen, with much love and gratitude, from @JaredPar's EditorUtils.</remarks>
	public static class VsMefContainerBuilder {
		#region Export Exclusion
		static readonly HashSet<string> excludedTypes = new HashSet<string> {
			// This uses IVsUIShell, which I haven't implemented, to show dialog boxes.
			// It also causes strange and fatal AccessViolations.
			"Microsoft.VisualStudio.Editor.Implementation.ExtensionErrorHandler",

			// This uses IOleComponentManager, which I don't know how to implement.
			"Microsoft.VisualStudio.Editor.Implementation.Intellisense.VsWpfKeyboardTrackingService",

			// This uses IWpfKeyboardTrackingService, and I don't want Code Lens anyway (yet?)
			"Microsoft.VisualStudio.Language.Intellisense.Implementation.CodeLensAdornmentCache",
			"Microsoft.VisualStudio.Language.Intellisense.Implementation.CodeLensInterLineAdornmentTaggerProvider",
		};
		///<summary>Prevents an exported type from being included in the created MEF container.</summary>
		///<remarks>Call this method if an exported type doesn't work outside Visual Studio.</remarks>
		public static void ExcludeExport(string fullTypeName) { excludedTypes.Add(fullTypeName); }

		///<summary>Creates a <see cref="ComposablePartCatalog"/> from the types in an assembly, excluding types that cause problems.</summary>
		public static ComposablePartCatalog GetFilteredCatalog(Assembly assembly) {
			return new TypeCatalog(assembly.GetTypes().Where(t => !excludedTypes.Contains(t.FullName)));
		}
		#endregion

		#region Catalog Setup
		private static readonly string[] EditorComponents = {
			// JaredPar: Core editor components
			"Microsoft.VisualStudio.Platform.VSEditor",

			// JaredPar: Not entirely sure why this is suddenly needed
			"Microsoft.VisualStudio.Text.Internal",

			// JaredPar: Must include this because several editor options are actually stored as exported information 
			// on this DLL.  Including most importantly, the tabsize information
			"Microsoft.VisualStudio.Text.Logic",

			// JaredPar: Include this DLL to get several more EditorOptions including WordWrapStyle
			"Microsoft.VisualStudio.Text.UI",

			// JaredPar: Include this DLL to get more EditorOptions values and the core editor
			"Microsoft.VisualStudio.Text.UI.Wpf",

			// SLaks: Needed for VisualStudioWaitIndicator & probably others
			"Microsoft.VisualStudio.Editor.Implementation",

			// SLaks: Needed for IVsHierarchyItemManager, used by peek providers
			"Microsoft.VisualStudio.Shell.TreeNavigation.HierarchyProvider"
		};

		static IEnumerable<ComposablePartCatalog> GetRoslynCatalogs() {
			if (VsLoader.RoslynAssemblyPath == null)
				return new ComposablePartCatalog[0];

			return Directory.EnumerateFiles(VsLoader.RoslynAssemblyPath, "Microsoft.CodeAnalysis*.dll")	// Leave out the . to catch Microsoft.CodeAnalysis.dll too
				.Select(p => GetFilteredCatalog(Assembly.LoadFile(p)))
				.Concat(new ComposablePartCatalog[] {
					GetFilteredCatalog(Assembly.Load("RoslynEditorHost")),
					new TypeCatalog(
						// IWaitIndicator is internal, so I have no choice but to use the existing
						// implementation. The rest of Microsoft.VisualStudio.LanguageServices.dll
						// exports lots of VS interop types that I don't want.
						Type.GetType("Microsoft.VisualStudio.LanguageServices.Implementation.Utilities.VisualStudioWaitIndicator, "
								   + "Microsoft.VisualStudio.LanguageServices")
					)
				});
		}
		#endregion

		// I need to specify a full name to load from the GAC.
		// The version is added by my AssemblyResolve handler.
		const string VsFullNameSuffix = ", Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL";

		///<summary>Gets all available ComposablePartCatalogs for the editor, including Roslyn if available.</summary>
		public static IEnumerable<ComposablePartCatalog> GetDefaultCatalogs() {
			return EditorComponents.Select(c => GetFilteredCatalog(Assembly.Load(c + VsFullNameSuffix)))
					.Concat(GetRoslynCatalogs())
					.Concat(new[] { GetFilteredCatalog(typeof(VsMefContainerBuilder).Assembly) });
		}

		///<summary>Creates and installs a CompositionContainer with all of the default catalogs.</summary>
		public static CompositionContainer CreateDefaultContainer(params ComposablePartCatalog[] additionalCatalogs) {
			return CreateDefaultContainer((IEnumerable<ComposablePartCatalog>)additionalCatalogs);
		}
		///<summary>Creates and installs a CompositionContainer with all of the default catalogs.</summary>
		public static CompositionContainer CreateDefaultContainer(IEnumerable<ComposablePartCatalog> additionalCatalogs) {
			var container = new CompositionContainer(new AggregateCatalog(GetDefaultCatalogs().Concat(additionalCatalogs)));
			InitialzeContainer(container);
			return container;
		}

		///<summary>
		/// Initializes an existing MEF container for use by the editor, and installs it into the global ServiceProvider.
		/// Editor factories will not work before this method is called.
		///</summary>
		public static void InitialzeContainer(CompositionContainer container) {
			// Based on Microsoft.VisualStudio.ComponentModelHost.ComponentModel.DefaultCompositionContainer.
			// By implementing SVsServiceProvider myself, I skip an unnecessary call to GetIUnknownForObject.
			container.ComposeExportedValue<SVsServiceProvider>(VsServiceProvider.Instance);

			// Needed because VsUndoHistoryRegistry tries to create IOleUndoManager from ILocalRegistry, which I presumably cannot do.
			container.ComposeExportedValue((ITextUndoHistoryRegistry)EditorUtils.EditorUtilsFactory.CreateBasicUndoHistoryRegistry());

			VsServiceProvider.Instance.SetMefContainer(container);

			// VS doesn't do this, but it's useful for my Roslyn adapter code.
			container.ComposeExportedValue(VsServiceProvider.Instance.ComponentModel);
		}
	}
}
