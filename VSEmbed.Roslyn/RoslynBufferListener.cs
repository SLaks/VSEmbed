using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel.Composition;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Host.Mef;
using Microsoft.CodeAnalysis.Text;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Utilities;

namespace VSEmbed.Roslyn {
	[Export(typeof(IWpfTextViewConnectionListener))]
	[ContentType("Roslyn Languages")]
	[TextViewRole(PredefinedTextViewRoles.Editable)]
	class RoslynBufferListener : IWpfTextViewConnectionListener {
		public SVsServiceProvider ExportProvider { get; private set; }

		[ImportingConstructor]
		public RoslynBufferListener(SVsServiceProvider exportProvider, RoslynSetup forceImport) {
			ExportProvider = exportProvider;
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

			var project = workspace.CurrentSolution
				.AddProject("Sample Project", "SampleProject", contentTypeLanguages[buffer.ContentType.DisplayName])
				.AddMetadataReferences(new[] { "mscorlib", "System", "System.Core", "System.Xml.Linq" }
					.Select(EditorWorkspace.CreateFrameworkReference)
				);

			workspace.TryApplyChanges(project.Solution);
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
