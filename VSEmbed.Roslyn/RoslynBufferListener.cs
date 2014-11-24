using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Collections.ObjectModel;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;
using Microsoft.CodeAnalysis.Host.Mef;
using Microsoft.CodeAnalysis.Text;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.LanguageServices;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Utilities;

namespace VSEmbed.Roslyn {
	[Export(typeof(IWpfTextViewConnectionListener))]
	[ContentType("Roslyn Languages")]
	[TextViewRole(PredefinedTextViewRoles.Editable)]
	public class RoslynBufferListener : IWpfTextViewConnectionListener {
		public SVsServiceProvider ExportProvider { get; private set; }

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

		[ImportingConstructor]
		public RoslynBufferListener(SVsServiceProvider exportProvider) {
			ExportProvider = exportProvider;
			var componentModel = (IComponentModel)ExportProvider.GetService(typeof(SComponentModel));

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
			diagnosticService.GetType()
				.GetField("workspaceAnalyzers", BindingFlags.NonPublic | BindingFlags.Instance)
				.SetValue(diagnosticService, new[] {
					"Microsoft.CodeAnalysis.Features.dll",
					"Microsoft.CodeAnalysis.EditorFeatures.dll",
					"Microsoft.CodeAnalysis.CSharp.Features.dll",
					"Microsoft.CodeAnalysis.CSharp.EditorFeatures.dll",
					"Microsoft.CodeAnalysis.VisualBasic.Features.dll",
					"Microsoft.CodeAnalysis.VisualBasic.EditorFeatures.dll",
				}.Select(name => new AnalyzerFileReference(Path.Combine(VsLoader.RoslynAssemblyPath, name), Assembly.LoadFile))
				 .ToImmutableArray<AnalyzerReference>());
		}

		static readonly Dictionary<string, string> contentTypeLanguages = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
			{ "CSharp", LanguageNames.CSharp },
			{ "Basic", LanguageNames.VisualBasic }
		};
		public async void SubjectBuffersConnected(IWpfTextView textView, ConnectionReason reason, Collection<ITextBuffer> subjectBuffers) {
			// Give the code that created the buffer a chance to attach its own workspace
			await Task.Yield();

			foreach (var buffer in subjectBuffers) {
				CreateWorkspace(buffer);
				buffer.ContentTypeChanged += (s, e) => {
					var workspace = buffer.GetWorkspace();
					if (workspace != null) {
						foreach (var document in buffer.GetRelatedDocuments()) {
							workspace.CloseDocument(document.Id);
						}
					}
					CreateWorkspace(buffer);
				};
			}
		}
		void CreateWorkspace(ITextBuffer buffer) {
			if (buffer.GetWorkspace() != null || !buffer.ContentType.IsOfType("Roslyn Languages"))
				return;
			var componentModel = (IComponentModel)ExportProvider.GetService(typeof(SComponentModel));
			var workspace = new EditorWorkspace(MefV1HostServices.Create(componentModel.DefaultExportProvider));

			var project = workspace.AddProject("Sample Project", contentTypeLanguages[buffer.ContentType.DisplayName]);
			workspace.TryApplyChanges(workspace.CurrentSolution.AddMetadataReferences(project.Id,
				new[] { "mscorlib", "System", "System.Core", "System.Xml.Linq" }.Select(workspace.CreateFrameworkReference)
			));
			workspace.CreateDocument(project.Id, buffer);
		}

		public void SubjectBuffersDisconnected(IWpfTextView textView, ConnectionReason reason, Collection<ITextBuffer> subjectBuffers) {
			foreach (var buffer in subjectBuffers) {
				foreach (var document in buffer.GetRelatedDocuments()) {
					buffer.GetWorkspace().CloseDocument(document.Id);
				}
			}
		}
	}
}
