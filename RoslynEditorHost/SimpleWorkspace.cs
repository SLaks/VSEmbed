using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;
using Microsoft.CodeAnalysis.Host;
using Microsoft.CodeAnalysis.Text;
using Microsoft.VisualStudio.Text;

namespace RoslynEditorHost {
	class SimpleWorkspace : Workspace {
		static readonly Type IWorkCoordinatorRegistrationService = Type.GetType("Microsoft.CodeAnalysis.SolutionCrawler.IWorkCoordinatorRegistrationService, Microsoft.CodeAnalysis.Features");

		public SimpleWorkspace(HostServices host) : base(host, "SimpleWorkspace") {
			var wcrService = typeof(HostWorkspaceServices)
				.GetMethod("GetService")
				.MakeGenericMethod(IWorkCoordinatorRegistrationService)
				.Invoke(Services, null);

			IWorkCoordinatorRegistrationService.GetMethod("Register").Invoke(wcrService, new[] { this });
		}
		public Project AddProject(string name, string language) {
			ProjectInfo projectInfo = ProjectInfo.Create(ProjectId.CreateNewId(null), VersionStamp.Create(), name, name, language);
			OnProjectAdded(projectInfo);
			return CurrentSolution.GetProject(projectInfo.Id);
		}
		public Document CreateDocument(ProjectId projectId, ITextBuffer buffer) {
			var id = DocumentId.CreateNewId(projectId);

			var docInfo = DocumentInfo.Create(id, "Sample Document",
				loader: TextLoader.From(buffer.AsTextContainer(), VersionStamp.Create()),
				sourceCodeKind: SourceCodeKind.Script
			);
			OnDocumentAdded(docInfo);
			OnDocumentOpened(id, buffer.AsTextContainer());
			buffer.Changed += delegate { OnDocumentContextUpdated(id); };
			return CurrentSolution.GetDocument(id);
		}

		protected override void AddMetadataReference(ProjectId projectId, MetadataReference metadataReference) {
			OnMetadataReferenceAdded(projectId, metadataReference);
		}
	}
}
