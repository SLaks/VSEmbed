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

namespace RoslynEditorHost {
	[Export(typeof(IWpfTextViewConnectionListener))]
	[ContentType("Roslyn Languages")]
	[TextViewRole(PredefinedTextViewRoles.Editable)]
	public class RoslynBufferListener : IWpfTextViewConnectionListener {
		public SVsServiceProvider ExportProvider { get; private set; }

		[ImportingConstructor]
		public RoslynBufferListener(SVsServiceProvider exportProvider) {
			ExportProvider = exportProvider;
			var componentModel = (IComponentModel)ExportProvider.GetService(typeof(SComponentModel));
			var container = (CompositionContainer)componentModel.DefaultExportProvider;

			// VisualStudioWaitIndicator imports VisualStudioWorkspace explicitly, and its ctor tries to use SQM.
			// Therefore, I hack together a barely-working instance and export it myself. This only works because
			// VisualStudioWaitIndicator fetches it from the ExportProvider instead of importing it normally.
			var vswType = Type.GetType("Microsoft.VisualStudio.LanguageServices.RoslynVisualStudioWorkspace, "
									 + "Microsoft.VisualStudio.LanguageServices.Implementation");
			var vsWorkspace = (VisualStudioWorkspace)FormatterServices.GetSafeUninitializedObject(vswType);

			// Initialize the base Workspace only (to set Services)
			typeof(Workspace).GetConstructors(BindingFlags.NonPublic | BindingFlags.Instance)[0]
				.Invoke(vsWorkspace, new object[] { MefHostServices.Create(container), "FakeWorkspace" });
			container.ComposeExportedValue<VisualStudioWorkspace>(vsWorkspace);
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
				var workspace = new SimpleWorkspace(MefHostServices.Create(componentModel.DefaultExportProvider));

				var project = workspace.AddProject("Sample Project", contentTypeLanguages[buffer.ContentType.DisplayName]);
				workspace.TryApplyChanges(workspace.CurrentSolution.AddMetadataReferences(project.Id, new[] {
					new MetadataFileReference(typeof(object).Assembly.Location, MetadataReferenceProperties.Assembly),
					new MetadataFileReference(typeof(Uri).Assembly.Location, MetadataReferenceProperties.Assembly),
					new MetadataFileReference(typeof(Enumerable).Assembly.Location, MetadataReferenceProperties.Assembly)
				}));
				workspace.CreateDocument(project.Id, buffer);
			}
		}

		public void SubjectBuffersDisconnected(IWpfTextView textView, ConnectionReason reason, Collection<ITextBuffer> subjectBuffers) {
		}
	}
}
