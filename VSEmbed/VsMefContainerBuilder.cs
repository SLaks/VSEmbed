using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.ComponentModel.Composition;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Media;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Composition;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.Text.Storage;
using Microsoft.VisualStudio.Utilities;

using MEFv1 = System.ComponentModel.Composition;
using MEFv3 = Microsoft.VisualStudio.Composition;

namespace VSEmbed {
	///<summary>Creates the MEF composition container used by the editor services.  This type is immutable</summary>
	/// <remarks>Stolen, with much love and gratitude, from @JaredPar's EditorUtils.</remarks>
	public abstract class VsMefContainerBuilder {
		#region Export Exclusion
		static readonly HashSet<string> excludedTypes = new HashSet<string> {
			// This uses IVsUIShell, which I haven't implemented, to show dialog boxes.
			// It also causes strange and fatal AccessViolations.
			"Microsoft.VisualStudio.Editor.Implementation.ExtensionErrorHandler",

			// This uses IOleUndoManager, which I don't have.  I replace it with @JaredPar's implementation.
			"Microsoft.VisualStudio.Editor.Implementation.Undo.VsUndoHistoryRegistry",

			// This uses IOleComponentManager, which I don't know how to implement.
			"Microsoft.VisualStudio.Editor.Implementation.Intellisense.VsWpfKeyboardTrackingService",

			// This uses IWpfKeyboardTrackingService, and I don't want Code Lens anyway (yet?)
			"Microsoft.VisualStudio.Language.Intellisense.Implementation.CodeLensAdornmentCache",
			"Microsoft.VisualStudio.Language.Intellisense.Implementation.CodeLensInterLineAdornmentTaggerProvider",

			// This uses COM services to try to read user settings, and I can't make that work.  I replace it with my own simple implementation.
			"Microsoft.VisualStudio.Editor.Implementation.DataStorageService",

			// I export my own direct SVsServiceProvider
			"Microsoft.VisualStudio.ComponentModelHost.VsComponentModelHostExporter"
		};
		///<summary>Prevents an exported type from being included in the created MEF container.</summary>
		///<remarks>Call this method if an exported type doesn't work outside Visual Studio.</remarks>
		public static void ExcludeExport(string fullTypeName) { excludedTypes.Add(fullTypeName); }
		#endregion

		///<summary>Creates a new builder, including a catalog with the types in a set of assemblies, excluding types that cause problems.</summary>
		public virtual VsMefContainerBuilder WithFilteredCatalogs(params Assembly[] assemblies) { return WithFilteredCatalogs((IEnumerable<Assembly>)assemblies); }
		///<summary>Creates a new builder, including a catalog with the types in a set of assemblies, excluding types that cause problems.</summary>
		public virtual VsMefContainerBuilder WithFilteredCatalogs(IEnumerable<Assembly> assemblies) {
			return WithCatalog(assemblies.SelectMany(a => a.GetTypes().Where(t => !excludedTypes.Contains(t.FullName))));
		}
		///<summary>Creates a new builder, including a catalog with the specified types.</summary>
		public abstract VsMefContainerBuilder WithCatalog(IEnumerable<Type> types);

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

		///<summary>Creates a new builder, including catalogs for the Roslyn language services.</summary>
		public VsMefContainerBuilder WithRoslynCatalogs() {
			if (VsLoader.RoslynAssemblyPath == null)
				return this;

			return WithFilteredCatalogs(
				// Only include assemblies that actually export useful MEF types.
				Directory.EnumerateFiles(VsLoader.RoslynAssemblyPath, "Microsoft.CodeAnalysis.*.dll")
					.Where(a => new[] { "Features", "Workspaces" }.Any(Path.GetFileName(a).Contains))
					.Select(Assembly.LoadFile))
				.WithFilteredCatalogs(Assembly.Load("VSEmbed.Roslyn"))
				.WithCatalog(new[] {
				// Roslyn formatter bug: https://roslyn.codeplex.com/workitem/382
				// IWaitIndicator is internal, so I have no choice but to use the existing
				// implementation. The rest of Microsoft.VisualStudio.LanguageServices.dll
				// exports lots of VS interop types that I don't want.
				Type.GetType("Microsoft.VisualStudio.LanguageServices.Implementation.Utilities.VisualStudioWaitIndicator, "
						   + "Microsoft.VisualStudio.LanguageServices"),
				// Enables rename (but breaks F12).
				Type.GetType("Microsoft.VisualStudio.LanguageServices.Implementation.VisualStudioDocumentNavigationServiceFactory, "
						   + "Microsoft.VisualStudio.LanguageServices"),
				// Provides error messages in quick fix previews.  This is in the VS layer
				// only because it uses VS icons, so I can use it as-is.
				// Removed in VS2015 Preview
				Type.GetType("Microsoft.VisualStudio.LanguageServices.Implementation.CodeFixPreview.CodeFixPreviewService, "
						   + "Microsoft.VisualStudio.LanguageServices"),
				// VS2015 Preview version of above type.
				Type.GetType("Microsoft.VisualStudio.LanguageServices.Implementation.PreviewPane.PreviewPaneService, "
						   + "Microsoft.VisualStudio.LanguageServices"),
				}.Where(t => t != null));

		}

		// I need to specify a full name to load from the GAC.
		// The version is added by my AssemblyResolve handler.
		const string VsFullNameSuffix = ", Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL";

		///<summary>Creates a new builder, including catalogs for the core editor services.</summary>
		public VsMefContainerBuilder WithEditorCatalogs() {
			return WithFilteredCatalogs(EditorComponents.Select(c => Assembly.Load(c + VsFullNameSuffix)))
				  .WithFilteredCatalogs(typeof(VsMefContainerBuilder).Assembly);
		}
		#endregion

		///<summary>Creates a builder prepopulated with the editor and Roslyn catalogs.</summary>
		public static VsMefContainerBuilder CreateDefault() {
			return Create().WithEditorCatalogs().WithRoslynCatalogs();
		}

		///<summary>Creates an empty builder with no catalogs.</summary>
		public static VsMefContainerBuilder Create() {
			if (Type.GetType("Microsoft.VisualStudio.Composition.CompositionConfiguration, Microsoft.VisualStudio.Composition") != null)
				return V3.Create();
			else
				return V1.Create();
		}

		class V1 : VsMefContainerBuilder {
			readonly ImmutableList<MEFv1.Primitives.ComposablePartCatalog> catalogs;

			public V1(ImmutableList<MEFv1.Primitives.ComposablePartCatalog> catalogs) {
				this.catalogs = catalogs;
			}

			internal new static VsMefContainerBuilder Create() {
				return new V1(ImmutableList<MEFv1.Primitives.ComposablePartCatalog>.Empty);
			}

			public override VsMefContainerBuilder WithCatalog(IEnumerable<Type> types) {
				return new V1(catalogs.Add(new MEFv1.Hosting.TypeCatalog(types)));
			}

			protected override IComponentModel BuildCore() {
				return new ComponentModel(new MEFv1.Hosting.CompositionContainer(new MEFv1.Hosting.AggregateCatalog(catalogs)));
			}

			class ComponentModel : IComponentModel {
				public ComponentModel(MEFv1.Hosting.CompositionContainer container) {
					DefaultCompositionService = container;
					DefaultExportProvider = container;
				}


				public MEFv1.ICompositionService DefaultCompositionService { get; private set; }
				public MEFv1.Hosting.ExportProvider DefaultExportProvider { get; private set; }

				public MEFv1.Primitives.ComposablePartCatalog DefaultCatalog { get { throw new NotSupportedException(); } }
				public MEFv1.Primitives.ComposablePartCatalog GetCatalog(string catalogName) { throw new NotSupportedException(); }

				public IEnumerable<T> GetExtensions<T>() where T : class { return DefaultExportProvider.GetExportedValues<T>(); }
				public T GetService<T>() where T : class { return DefaultExportProvider.GetExportedValue<T>(); }
			}
		}

		class V3 : VsMefContainerBuilder {
			internal static new VsMefContainerBuilder Create() {
				return new V3(MEFv3.ComposableCatalog.Create())
					// Needed for ExportMetadataViewInterfaceEmitProxy to support editor metadata types.
					.WithFilteredCatalogs(Assembly.Load("Microsoft.VisualStudio.Composition.Configuration")); ;
			}
			readonly MEFv3.ComposableCatalog catalog;
			private V3(MEFv3.ComposableCatalog catalog) {
				this.catalog = catalog;
			}


			static readonly MEFv3.PartDiscovery partDiscovery = MEFv3.PartDiscovery.Combine(
				new MEFv3.AttributedPartDiscovery { IsNonPublicSupported = true },
				new MEFv3.AttributedPartDiscoveryV1()
			);
			public override VsMefContainerBuilder WithCatalog(IEnumerable<Type> types) {
				// Consumers are expected to build their MEF catalogs before setting
				// up the UI thread, so this should not create async deadlocks under
				// normal usage.
				return new V3(catalog.WithParts(
					partDiscovery.CreatePartsAsync(types)
						.GetAwaiter()
						.GetResult()
						.ThrowOnErrors()
				));
			}

			protected override IComponentModel BuildCore() {
				var exportProvider = MEFv3.RuntimeComposition
					.CreateRuntimeComposition(MEFv3.CompositionConfiguration.Create(catalog).ThrowOnErrors())
					.CreateExportProviderFactory()
					.CreateExportProvider();
				return new ComponentModel(exportProvider);
			}
			class ComponentModel : IComponentModel {
				public readonly MEFv3.ExportProvider ExportProvider;

				public ComponentModel(MEFv3.ExportProvider exportProvider) {
					this.ExportProvider = exportProvider;
				}


				public MEFv1.ICompositionService DefaultCompositionService {
					get { return ExportProvider.GetExport<MEFv1.ICompositionService>().Value; }
				}

				public MEFv1.Hosting.ExportProvider DefaultExportProvider {
					get { return ExportProvider.AsExportProvider(); }
				}

				public MEFv1.Primitives.ComposablePartCatalog DefaultCatalog { get { throw new NotSupportedException(); } }
				public MEFv1.Primitives.ComposablePartCatalog GetCatalog(string catalogName) { throw new NotSupportedException(); }

				public IEnumerable<T> GetExtensions<T>() where T : class { return ExportProvider.GetExportedValues<T>(); }
				public T GetService<T>() where T : class { return ExportProvider.GetExportedValue<T>(); }
			}
		}

		///<summary>
		/// Creates a MEF container from this builder instance, and installs it into the global ServiceProvider.
		/// Editor factories will not work before this method is called.
		///</summary>
		public IComponentModel Build() {
			var container = BuildCore();
			VsServiceProvider.Instance.SetMefContainer(container);
			return container;
		}

		///<summary>Builds the described MEF container and wraps it in a Visual Studio <see cref="IComponentModel"/> implementation.</summary>
		protected abstract IComponentModel BuildCore();

	}
}
