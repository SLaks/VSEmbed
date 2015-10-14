using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Host;
using Microsoft.CodeAnalysis.Host.Mef;
using Microsoft.CodeAnalysis.Text;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;

namespace VSEmbed.Roslyn {
	///<summary>A Roslyn Workspace that contains documents linked to <see cref="ITextBuffer"/>s.</summary>
	public class EditorWorkspace : Workspace {
		// TODO: Add an optional parameter to pass changes through to an existing MSBuildWorkspace

		static readonly Type ISolutionCrawlerRegistrationService = Type.GetType("Microsoft.CodeAnalysis.SolutionCrawler.ISolutionCrawlerRegistrationService, Microsoft.CodeAnalysis.Features");
		static readonly Type IDocumentTrackingService = Type.GetType("Microsoft.CodeAnalysis.IDocumentTrackingService, Microsoft.CodeAnalysis.Features");

		readonly Dictionary<DocumentId, ITextBuffer> documentBuffers = new Dictionary<DocumentId, ITextBuffer>();
		///<summary>Creates an <see cref="EditorWorkspace"/> powered by the specified MEF host services.</summary>
		public EditorWorkspace(HostServices host) : base(host, WorkspaceKind.Host) {
			(host as MefV1HostServices)?.GetExports<RoslynSetup>().Single().Value.ToString();

			ISolutionCrawlerRegistrationService.GetMethod("Register")
				.Invoke(GetInternalService(ISolutionCrawlerRegistrationService), new[] { this });

			// TODO: http://source.roslyn.codeplex.com/#Microsoft.CodeAnalysis.EditorFeatures/Implementation/Workspaces/ProjectCacheService.cs,63?
		}

		///<summary>Gets a non-public <see cref="IWorkspaceService"/> from this instance.</summary>
		private object GetInternalService(Type interfaceType) {
			return typeof(HostWorkspaceServices)
				.GetMethod("GetService")
				.MakeGenericMethod(interfaceType)
				.Invoke(Services, null);
		}

		DocumentId activeDocumentId;

		class FakeVsWindowFrame : IVsWindowFrame, IVsWindowFrame2 {
			public int ActivateOwnerDockedWindow() {
				throw new NotImplementedException();
			}

			public int Advise(IVsWindowFrameNotify pNotify, out uint pdwCookie) {
				pdwCookie = 0;
				return 0;
			}

			public int CloseFrame(uint grfSaveOptions) {
				throw new NotImplementedException();
			}

			public int GetFramePos(VSSETFRAMEPOS[] pdwSFP, out Guid pguidRelativeTo, out int px, out int py, out int pcx, out int pcy) {
				throw new NotImplementedException();
			}

			public int GetGuidProperty(int propid, out Guid pguid) {
				throw new NotImplementedException();
			}

			public int GetProperty(int propid, out object pvar) {
				throw new NotImplementedException();
			}

			public int Hide() {
				throw new NotImplementedException();
			}

			public int IsOnScreen(out int pfOnScreen) {
				throw new NotImplementedException();
			}

			public int IsVisible() {
				throw new NotImplementedException();
			}

			public int QueryViewInterface(ref Guid riid, out IntPtr ppv) {
				throw new NotImplementedException();
			}

			public int SetFramePos(VSSETFRAMEPOS dwSFP, ref Guid rguidRelativeTo, int x, int y, int cx, int cy) {
				throw new NotImplementedException();
			}

			public int SetGuidProperty(int propid, ref Guid rguid) {
				throw new NotImplementedException();
			}

			public int SetProperty(int propid, object var) {
				throw new NotImplementedException();
			}

			public int Show() {
				throw new NotImplementedException();
			}

			public int ShowNoActivate() {
				throw new NotImplementedException();
			}

			public int Unadvise(uint dwCookie) {
				throw new NotImplementedException();
			}
		}

		// TODO: Let callers pick a framework version
		static readonly string referenceAssemblyPath = Path.Combine(
			Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
			@"Reference Assemblies\Microsoft\Framework\.NETFramework\v4.5"
		);

		static readonly Type VisualStudioDocumentTrackingService = Type.GetType("Microsoft.VisualStudio.LanguageServices.Implementation.VisualStudioDocumentTrackingService, Microsoft.VisualStudio.LanguageServices");
		static readonly Type FrameListener = Type.GetType("Microsoft.VisualStudio.LanguageServices.Implementation.VisualStudioDocumentTrackingService+FrameListener, Microsoft.VisualStudio.LanguageServices");
		static readonly FieldInfo activeFrameField = VisualStudioDocumentTrackingService.GetField("_activeFrame", BindingFlags.NonPublic | BindingFlags.Instance);
		static readonly FieldInfo visibleFramesField = VisualStudioDocumentTrackingService.GetField("_visibleFrames", BindingFlags.NonPublic | BindingFlags.Instance);
		static readonly MethodInfo CreateImmutableList = new Func<int, ImmutableList<int>>(ImmutableList.Create)
			.Method
			.GetGenericMethodDefinition()
			.MakeGenericMethod(FrameListener);

		///<summary>Gets or sets the document that the user is editing.  Set this property to speed up semantic processing (live error checking).</summary>
		public DocumentId ActiveDocumentId {
			get { return activeDocumentId; }
			set {
				if (ActiveDocumentId == value) return;
				activeDocumentId = value;
				// This is completely coupled to http://source.roslyn.codeplex.com/#Microsoft.VisualStudio.LanguageServices/Implementation/Workspace/VisualStudioDocumentTrackingService.cs
				// This is necessary to force edits to be processed by the WorkCoordinator.HighPriorityProcessor.
				var docTracker = GetInternalService(IDocumentTrackingService);
				if (value == null)
					activeFrameField.SetValue(docTracker, null);
				else {
					var frame = new FakeVsWindowFrame();
					activeFrameField.SetValue(docTracker, frame);

					var listener = Activator.CreateInstance(FrameListener, docTracker, frame, value);
					visibleFramesField.SetValue(docTracker, CreateImmutableList.Invoke(null, new[] { listener }));
				}
			}
		}

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
			ActiveDocumentId = documentId;
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
