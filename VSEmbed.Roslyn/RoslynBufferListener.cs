using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Host.Mef;
using Microsoft.CodeAnalysis.Text;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.LanguageServices;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Utilities;

namespace VSEmbed.Roslyn {
	///<summary>An [Export] attribute that can export an inaccessible interface.</summary>
	[AttributeUsage(AttributeTargets.Class | AttributeTargets.Method | AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = true, Inherited = false)]
	sealed class HackyExportAttribute : ExportAttribute {
		public HackyExportAttribute(string qualifiedTypeName) : base(Type.GetType(qualifiedTypeName)) { }
	}

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

		// Roslyn loads analyzers from DLL filenames that come from this VS service,
		// which gets DLL location VS's extension manager service.  I inject a faked
		// instance which just provides the DLLs containing Roslyn's own analyzers.
		//[HackyExport("Microsoft.CodeAnalysis.Diagnostics.IWorkspaceDiagnosticAnalyzerProviderService, Microsoft.CodeAnalysis.Features")]
		static readonly object analyzerProvider = FormatterServices.GetSafeUninitializedObject(
			Type.GetType("Microsoft.VisualStudio.LanguageServices.Implementation.Diagnostics.VisualStudioWorkspaceDiagnosticAnalyzerProviderService, "
					   + "Microsoft.VisualStudio.LanguageServices"));

		[ImportingConstructor]
		public RoslynBufferListener(SVsServiceProvider exportProvider) {
			ExportProvider = exportProvider;
			var componentModel = (IComponentModel)ExportProvider.GetService(typeof(SComponentModel));

			// Initialize the base Workspace only (to set Services)
			// The MefV1HostServices call breaks compatibility with
			// older Dev14 CTPs; that could be fixed by reflection.
			typeof(Workspace).GetConstructors(BindingFlags.NonPublic | BindingFlags.Instance)[0]
				.Invoke(vsWorkspace, new object[] { MefV1HostServices.Create(componentModel.DefaultExportProvider), "FakeWorkspace" });

			analyzerProvider.GetType().GetField("workspaceAnalyzerAssemblies", BindingFlags.NonPublic | BindingFlags.Instance)
				.SetValue(analyzerProvider, new string[] {
					"Microsoft.CodeAnalysis.Features",
					"Microsoft.CodeAnalysis.EditorFeatures",
					"Microsoft.CodeAnalysis.CSharp.Features",
					"Microsoft.CodeAnalysis.CSharp.EditorFeatures",
					"Microsoft.CodeAnalysis.VisualBasic.Features",
					"Microsoft.CodeAnalysis.VisualBasic.EditorFeatures",
				});
		}

		static readonly Dictionary<string, string> contentTypeLanguages = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
			{ "CSharp", LanguageNames.CSharp },
			{ "Basic", LanguageNames.VisualBasic }
		};
		public async void SubjectBuffersConnected(IWpfTextView textView, ConnectionReason reason, Collection<ITextBuffer> subjectBuffers) {
			// Give the code that created the buffer a chance to attach its own workspace
			await Task.Yield();
			foreach (var buffer in subjectBuffers) {
				if (buffer.GetWorkspace() != null)
					continue;
				var componentModel = (IComponentModel)ExportProvider.GetService(typeof(SComponentModel));
				var workspace = new EditorWorkspace(MefV1HostServices.Create(componentModel.DefaultExportProvider));

				var project = workspace.AddProject("Sample Project", contentTypeLanguages[buffer.ContentType.DisplayName]);
				workspace.TryApplyChanges(workspace.CurrentSolution.AddMetadataReferences(project.Id,
					new[] { "mscorlib", "System", "System.Core", "System.Xml.Linq" }.Select(workspace.CreateFrameworkReference)
				));
				workspace.CreateDocument(project.Id, buffer);
			}
		}

		public void SubjectBuffersDisconnected(IWpfTextView textView, ConnectionReason reason, Collection<ITextBuffer> subjectBuffers) {
		}
	}
}
