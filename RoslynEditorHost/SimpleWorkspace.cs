using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Host;
using Microsoft.CodeAnalysis.Text;

namespace RoslynEditorHost {
	class SimpleWorkspace : Workspace {
		public SimpleWorkspace(HostServices host) : base(host, "SimpleWorkspace") { }
		public Project AddProject(string name, string language) {
			ProjectInfo projectInfo = ProjectInfo.Create(ProjectId.CreateNewId(null), VersionStamp.Create(), name, name, language);
			OnProjectAdded(projectInfo);
			return CurrentSolution.GetProject(projectInfo.Id);
		}
		public Document CreateDocument(ProjectId projectId, SourceTextContainer container) {
			var id = DocumentId.CreateNewId(projectId);

			var docInfo = DocumentInfo.Create(id, "Sample Document",
				loader: TextLoader.From(container, VersionStamp.Create()),
				sourceCodeKind: SourceCodeKind.Script
			);
			OnDocumentAdded(docInfo);
			OnDocumentOpened(id, container);
			return CurrentSolution.GetDocument(id);
		}
	}
}
