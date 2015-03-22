using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;
using Microsoft.CodeAnalysis.Host;
using Microsoft.CodeAnalysis.MSBuild;
using Microsoft.CodeAnalysis.Text;
using Microsoft.VisualStudio.Text;

namespace VSEmbed.Roslyn {
	///<summary>A Roslyn Workspace that contains documents linked to <see cref="ITextBuffer"/>s.</summary>
	public class EditorWorkspace : Workspace {
		// TODO: Add an optional parameter to pass changes through to an existing MSBuildWorkspace

		static readonly Type IWorkCoordinatorRegistrationService = Type.GetType("Microsoft.CodeAnalysis.SolutionCrawler.IWorkCoordinatorRegistrationService, Microsoft.CodeAnalysis.Features");

		readonly Dictionary<DocumentId, ITextBuffer> documentBuffers = new Dictionary<DocumentId, ITextBuffer>();
		///<summary>Creates an <see cref="EditorWorkspace"/> powered by the specified MEF host services.</summary>
		public EditorWorkspace(HostServices host) : base(host, WorkspaceKind.Host) {
			var wcrService = typeof(HostWorkspaceServices)
				.GetMethod("GetService")
				.MakeGenericMethod(IWorkCoordinatorRegistrationService)
				.Invoke(Services, null);

			IWorkCoordinatorRegistrationService.GetMethod("Register").Invoke(wcrService, new[] { this });
		}

		// TODO: Let callers pick a framework version
		static readonly string referenceAssemblyPath = Path.Combine(
			Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
			@"Reference Assemblies\Microsoft\Framework\.NETFramework\v4.5"
		);

		///<summary>Creates a <see cref="MetadataReference"/> to a BCL assembly, with XML documentation.</summary>
		public static MetadataReference CreateFrameworkReference(string assemblyName) {
			return MetadataReference.CreateFromFile(
				Path.Combine(referenceAssemblyPath, assemblyName + ".dll"),
				MetadataReferenceProperties.Assembly,
				new XmlDocumentationProvider(Path.Combine(referenceAssemblyPath, assemblyName + ".xml"))
			);
		}

		///<summary>Creates a new document linked to an existing text buffer.</summary>
		public DocumentId CreateDocument(ProjectId projectId, ITextBuffer buffer, string debugName = null) {
			var id = DocumentId.CreateNewId(projectId, debugName);

			TryApplyChanges(CurrentSolution.AddDocument(id, debugName ?? "Sample Document", TextLoader.From(buffer.AsTextContainer(), VersionStamp.Create())));
			OpenDocument(id, buffer);
			return id;
		}

		///<summary>Links an existing <see cref="Document"/> to an <see cref="ITextBuffer"/>, synchronizing their contents.</summary>
		public void OpenDocument(DocumentId documentId, ITextBuffer buffer) {
			documentBuffers.Add(documentId, buffer);
			OnDocumentOpened(documentId, buffer.AsTextContainer());
			buffer.Changed += delegate { OnDocumentContextUpdated(documentId); };
		}

		///<summary>Unlinks an opened <see cref="Document"/> from its <see cref="ITextBuffer"/>.</summary>
		public override void CloseDocument(DocumentId documentId) {
			var document = CurrentSolution.GetDocument(documentId);
			OnDocumentClosed(documentId, TextLoader.From(TextAndVersion.Create(document.GetTextAsync().Result, document.GetTextVersionAsync().Result)));
			documentBuffers.Remove(documentId);
		}

		///<summary>Applies document text changes to documents backed by <see cref="ITextBuffer"/>s.</summary>
		protected override void ApplyDocumentTextChanged(DocumentId id, SourceText text) {
			ITextBuffer buffer;
			if (documentBuffers.TryGetValue(id, out buffer))
				UpdateText(text, buffer, EditOptions.DefaultMinimalChange);
			else
				base.ApplyDocumentTextChanged(id, text);
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

		///<summary><see cref="EditorWorkspace"/> can apply any kind of change.</summary>
		public override bool CanApplyChange(ApplyChangesKind feature) {
			return true;
		}
	}
}
