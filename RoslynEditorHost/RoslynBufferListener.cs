using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Host.Mef;
using Microsoft.CodeAnalysis.Text;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Utilities;

namespace RoslynEditorHost {
	[Export(typeof(IWpfTextViewConnectionListener))]
	[ContentType("Roslyn Languages")]
	[TextViewRole(PredefinedTextViewRoles.Editable)]
	public class RoslynBufferListener : IWpfTextViewConnectionListener {
		[Import]
		public SVsServiceProvider ExportProvider { get; set; }

		static readonly Dictionary<string, string> contentTypeLanguages = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
			{ "CSharp", "C#" },
			{ "Basic", "Visual Basic" }
		};
		public void SubjectBuffersConnected(IWpfTextView textView, ConnectionReason reason, Collection<ITextBuffer> subjectBuffers) {
			foreach (var buffer in subjectBuffers) {
				if (buffer.GetWorkspace() != null)
					continue;
				var componentModel = (IComponentModel)ExportProvider.GetService(typeof(SComponentModel));
				var workspace = new SimpleWorkspace(MefHostServices.Create(componentModel.DefaultExportProvider));

				var project = workspace.AddProject("Sample Project", contentTypeLanguages[buffer.ContentType.DisplayName]);
				workspace.TryApplyChanges(workspace.CurrentSolution.AddMetadataReferences(project.Id, new[] {
					new MetadataFileReference(typeof(object).Assembly.Location, MetadataReferenceProperties.Assembly),
					new MetadataFileReference(typeof(Uri).Assembly.Location, MetadataReferenceProperties.Assembly),
					new MetadataFileReference(typeof(Enumerable).Assembly.Location, MetadataReferenceProperties.Assembly)
				}));
				workspace.CreateDocument(project.Id, buffer.AsTextContainer());
			}
		}

		public void SubjectBuffersDisconnected(IWpfTextView textView, ConnectionReason reason, Collection<ITextBuffer> subjectBuffers) {
		}
	}
}
