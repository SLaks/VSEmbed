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

namespace VSEmbed.Roslyn {
	///<summary>A Roslyn Workspace that contains documents linked to ITextBuffers.</summary>
	public class EditorWorkspace : Workspace {
		// TODO: Add an optional parameter to pass changes through to an existing MSBuildWorkspace

		static readonly Type IWorkCoordinatorRegistrationService = Type.GetType("Microsoft.CodeAnalysis.SolutionCrawler.IWorkCoordinatorRegistrationService, Microsoft.CodeAnalysis.Features");

		readonly Dictionary<DocumentId, ITextBuffer> documentBuffers = new Dictionary<DocumentId, ITextBuffer>();
		public EditorWorkspace(HostServices host) : base(host, WorkspaceKind.Host) {
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

		///<summary>Creates a new document linked to an existing text buffer.</summary>
		public Document CreateDocument(ProjectId projectId, ITextBuffer buffer) {
			var id = DocumentId.CreateNewId(projectId);
			documentBuffers.Add(id, buffer);

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
		protected override void ChangedDocumentText(DocumentId id, SourceText text) {
			OnDocumentTextChanged(id, text, PreservationMode.PreserveValue);
			UpdateText(text, documentBuffers[id], EditOptions.DefaultMinimalChange);
		}

		// Stolen from Microsoft.VisualStudio.LanguageServices.Implementation.ProjectSystem.DocumentProvider.StandardTextDocument
		private static void UpdateText(SourceText newText, ITextBuffer buffer, EditOptions options) {
			using (ITextEdit textEdit = buffer.CreateEdit(options, null, null)) {
				SourceText oldText = buffer.CurrentSnapshot.AsText();
				foreach (TextChange current in newText.GetTextChanges(oldText)) {
					textEdit.Replace(current.Span.Start, current.Span.Length, current.NewText);
				}
				textEdit.Apply();
			}
		}

		public override bool CanApplyChange(ApplyChangesKind feature) {
			switch (feature) {
				case ApplyChangesKind.AddMetadataReference:
				case ApplyChangesKind.RemoveMetadataReference:
				case ApplyChangesKind.ChangeDocument:
					return true;
				case ApplyChangesKind.AddProject:
				case ApplyChangesKind.RemoveProject:
				case ApplyChangesKind.AddProjectReference:
				case ApplyChangesKind.RemoveProjectReference:
				case ApplyChangesKind.AddDocument:
				case ApplyChangesKind.RemoveDocument:
				default:
					return false;
			}
		}
	}
}
