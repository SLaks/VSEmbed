using System;
using System.Collections.Immutable;
using System.ComponentModel.Composition;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;
using Microsoft.CodeAnalysis.Host.Mef;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.LanguageServices;
using Microsoft.VisualStudio.Shell;

namespace VSEmbed.Roslyn {
	///<summary>
	/// Import this class from MEF to ensure that all Roslyn hooks have been initialized.  
	/// This must be imported before any workspaces from the MEF container are used.  
	/// <see cref="EditorWorkspace"/> will import this automatically.
	///</summary>
	[Export]
	public class RoslynSetup {
		// VisualStudioWaitIndicator imports VisualStudioWorkspace explicitly, and its ctor tries to use SQM.
		// Therefore, I hack together a barely-working instance and export it myself.
		static readonly Type vswType = Type.GetType("Microsoft.VisualStudio.LanguageServices.RoslynVisualStudioWorkspace, "
												  + "Microsoft.VisualStudio.LanguageServices.Implementation");
		// I create an uninitialized instance and export it to MEF, then initialize it with the MEF container
		// later, once I have access to the container. The VS code solves this by exporting a class that gets
		// the ComponentModel from the ServiceProvider and passes it to its base ctor.
		[Export]
		static readonly VisualStudioWorkspace vsWorkspace =
			(VisualStudioWorkspace)FormatterServices.GetSafeUninitializedObject(vswType);



		///<summary>Sets up all required Roslyn hooks.  This is called by MEF.</summary>
		[ImportingConstructor]
		public RoslynSetup(SVsServiceProvider exportProvider) {
			var componentModel = (IComponentModel)exportProvider.GetService(typeof(SComponentModel));

			// Initialize the base Workspace only (to set Services)
			// The MefV1HostServices call breaks compatibility with
			// older Dev14 CTPs; that could be fixed by reflection.
			typeof(Workspace).GetConstructors(BindingFlags.NonPublic | BindingFlags.Instance)[0]
				.Invoke(vsWorkspace, new object[] { MefV1HostServices.Create(componentModel.DefaultExportProvider), "FakeWorkspace" });


			var diagnosticService = componentModel.DefaultExportProvider
				.GetExport<object>("Microsoft.CodeAnalysis.Diagnostics.IDiagnosticAnalyzerService").Value;

			// Roslyn loads analyzers from DLL filenames that come from the VS-layer
			// IWorkspaceDiagnosticAnalyzerProviderService. This uses internal types
			// which I cannot provide. Instead, I inject the standard analyzers into
			// DiagnosticAnalyzerService myself, after it's created.
			var analyzerManager = diagnosticService.GetType()
				.GetField("_hostAnalyzerManager", BindingFlags.NonPublic | BindingFlags.Instance)
				.GetValue(diagnosticService);
			analyzerManager.GetType()
				.GetField("_hostAnalyzerReferencesMap", BindingFlags.NonPublic | BindingFlags.Instance)
				.SetValue(analyzerManager, new[] {
					"Microsoft.CodeAnalysis.Features.dll",
					"Microsoft.CodeAnalysis.EditorFeatures.dll",
					"Microsoft.CodeAnalysis.CSharp.dll",
					"Microsoft.CodeAnalysis.CSharp.Features.dll",
					"Microsoft.CodeAnalysis.CSharp.EditorFeatures.dll",
					"Microsoft.CodeAnalysis.VisualBasic.dll",
					"Microsoft.CodeAnalysis.VisualBasic.Features.dll",
					"Microsoft.CodeAnalysis.VisualBasic.EditorFeatures.dll",
				}.Select(name => new AnalyzerFileReference(
					Path.Combine(VsLoader.RoslynAssemblyPath, name),
					new AnalyzerLoader()
				))
				 .ToImmutableDictionary<AnalyzerReference, object>(a => a.Id));
			// Based on HostAnalyzerManager.CreateAnalyzerReferencesMap

			var packageType = Type.GetType("Microsoft.VisualStudio.LanguageServices.Setup.RoslynPackage, Microsoft.VisualStudio.LanguageServices");
			var package = Activator.CreateInstance(packageType, nonPublic: true);
			// Bind Roslyn UI to VS theme colors
			packageType.GetMethod("InitializeColors", BindingFlags.Instance | BindingFlags.NonPublic)
					   .Invoke(package, null);
		}
		class AnalyzerLoader : IAnalyzerAssemblyLoader {
			public void AddDependencyLocation(string fullPath) { }

			public Assembly LoadFromPath(string fullPath) => Assembly.Load(AssemblyName.GetAssemblyName(fullPath));
		}
	}
}
