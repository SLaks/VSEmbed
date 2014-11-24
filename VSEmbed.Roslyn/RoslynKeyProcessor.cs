using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Windows.Threading;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Services;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Utilities;
using VSEmbed.Editor;


namespace VSEmbed.Roslyn {

	// The Roslyn commanding system is built around ICommandHandlerService, which intercepts and
	// handles every command.  Unfortunately, the entire system is internal.  I steal the entry-
	// points from AbstractOleCommandTarget and invoke them directly using Reflection. Note that
	// this is rather brittle.
	class RoslynKeyProcessor : ChainedKeyProcessor {
		// This delegate matches the signature of the Execute*() methods in AbstractOleCommandTarget
		delegate void CommandExecutor(ITextBuffer subjectBuffer, IContentType contentType, Action executeNextCommandTarget);
		// This delegate matches ExecuteUndo and ExecuteRedo
		delegate void IntCommandExecutor(ITextBuffer subjectBuffer, IContentType contentType, Action executeNextCommandTarget, int count);
		readonly IWpfTextView wpfTextView;
		readonly object innerCommandTarget;

		static readonly Type packageType = Type.GetType("Microsoft.VisualStudio.LanguageServices.CSharp.LanguageService.CSharpPackage, "
													  + "Microsoft.VisualStudio.LanguageServices.CSharp");
		static readonly Type languageServiceType = Type.GetType("Microsoft.VisualStudio.LanguageServices.CSharp.LanguageService.CSharpLanguageService, "
															  + "Microsoft.VisualStudio.LanguageServices.CSharp");
		// The generic parameters aren't actually used, so there is nothing wrong with always using C#.
		// The methods I call are on the non-generic abstract base class anyway.
		static readonly Type oleCommandTargetType = Type.GetType("Microsoft.VisualStudio.LanguageServices.Implementation.StandaloneCommandFilter`3, "
															   + "Microsoft.VisualStudio.LanguageServices")
			.MakeGenericType(packageType, languageServiceType,
				Type.GetType("Microsoft.VisualStudio.LanguageServices.CSharp.ProjectSystemShim.CSharpProject, Microsoft.VisualStudio.LanguageServices.CSharp")
			);

		public RoslynKeyProcessor(IWpfTextView wpfTextView, IComponentModel mef) {
			this.wpfTextView = wpfTextView;
			innerCommandTarget = CreateInstanceNonPublic(oleCommandTargetType,
				CreateInstanceNonPublic(languageServiceType, Activator.CreateInstance(packageType, true)),	// languageService
				wpfTextView,										// wpfTextView
				mef.DefaultExportProvider.GetExport<object>("Microsoft.CodeAnalysis.Editor.ICommandHandlerServiceFactory").Value,			// commandHandlerServiceFactory
				null,												// featureOptionsService (not used)
				mef.GetService<IVsEditorAdaptersFactoryService>()	// editorAdaptersFactoryService
			);

			AddShortcuts();
		}
		static object CreateInstanceNonPublic(Type type, params object[] args) {
			return Activator.CreateInstance(type, BindingFlags.NonPublic | BindingFlags.Instance, null, args, null);
		}

		readonly Dictionary<Tuple<ModifierKeys, Key>, CommandExecutor> shortcuts = new Dictionary<Tuple<ModifierKeys, Key>, CommandExecutor>();

		protected void AddCommand(Key key, string methodName) {
			AddCommand(ModifierKeys.None, key, methodName);
		}
		protected void AddShiftCommand(Key key, string methodName) {
			AddCommand(ModifierKeys.Shift, key, methodName);
		}
		protected void AddControlCommand(Key key, string methodName) {
			AddCommand(ModifierKeys.Control, key, methodName);
		}
		protected void AddControlShiftCommand(Key key, string methodName) {
			AddCommand(ModifierKeys.Control | ModifierKeys.Shift, key, methodName);
		}
		protected void AddAltShiftCommand(Key key, string methodName) {
			AddCommand(ModifierKeys.Alt | ModifierKeys.Shift, key, methodName);
		}
		protected void AddCommand(ModifierKeys modifiers, Key key, string methodName) {
			var method = (CommandExecutor)Delegate.CreateDelegate(typeof(CommandExecutor), innerCommandTarget, methodName);
			shortcuts.Add(new Tuple<ModifierKeys, Key>(modifiers, key), method);
		}
		protected void AddIntCommand(ModifierKeys modifiers, Key key, string methodName, int count) {
			var method = (IntCommandExecutor)Delegate.CreateDelegate(typeof(IntCommandExecutor), innerCommandTarget, methodName);
			shortcuts.Add(new Tuple<ModifierKeys, Key>(modifiers, key), (s, c, n) => method(s, c, n, count));
		}

		public override void KeyDown(KeyEventArgs args, ITextBuffer targetBuffer, Action next) {
			CommandExecutor method;
			if (shortcuts.TryGetValue(Tuple.Create(args.KeyboardDevice.Modifiers, args.Key), out method))
				method(targetBuffer, targetBuffer.ContentType, next);
			else
				next();
		}

		// ExecuteTypeCharacter() takes a COM variant pointer.
		// Instead of messing with Marshal and variants, I use
		// more Reflection to bypass that method and call into
		// the ICommandHandlerService directly.
		static readonly PropertyInfo currentHandlersProperty = oleCommandTargetType.GetProperty("CurrentHandlers", BindingFlags.NonPublic | BindingFlags.Instance);
		static readonly Type TypeCharCommandArgs = Type.GetType("Microsoft.CodeAnalysis.Editor.Commands.TypeCharCommandArgs, Microsoft.CodeAnalysis.EditorFeatures");
		static readonly MethodInfo executeMethod = currentHandlersProperty.PropertyType.GetMethod("Execute").MakeGenericMethod(TypeCharCommandArgs);

		public override void TextInput(TextCompositionEventArgs args, ITextBuffer targetBuffer, Action next) {
			if (args.Text.Length == 0) {
				next();
				return;
			}
			if (args.Text.Length > 1)
				throw new InvalidOperationException("I cannot properly forward multi-character text input");
			var commandArgs = Activator.CreateInstance(TypeCharCommandArgs, wpfTextView, targetBuffer, args.Text[0]);

			var currentHandlers = currentHandlersProperty.GetValue(innerCommandTarget);
			executeMethod.Invoke(currentHandlers, new object[] { targetBuffer.ContentType, commandArgs, next });
		}

		#region Shortcuts
		void AddShortcuts() {
			#region Cursor Movement
			AddCommand(Key.Up, "ExecuteUp");
			AddCommand(Key.Down, "ExecuteDown");
			AddCommand(Key.PageUp, "ExecutePageUp");
			AddCommand(Key.PageDown, "ExecutePageDown");

			AddCommand(Key.Home, "ExecuteLineStart");
			AddCommand(Key.End, "ExecuteLineEnd");
			AddShiftCommand(Key.Home, "ExecuteLineStartExtend");
			AddShiftCommand(Key.End, "ExecuteLineEndExtend");

			AddControlCommand(Key.Home, "ExecuteDocumentStart");
			AddControlCommand(Key.End, "ExecuteDocumentEnd");

			AddControlCommand(Key.A, "ExecuteSelectAll");
			#endregion

			AddCommand(Key.F12, "ExecuteGotoDefinition");
			AddCommand(Key.F2, "ExecuteRename");
			AddCommand(Key.Escape, "ExecuteCancel");
			AddControlShiftCommand(Key.Space, "ExecuteParameterInfo");
			AddControlCommand(Key.Space, "ExecuteCommitUniqueCompletionItem");

			AddControlShiftCommand(Key.Down, "ExecutePreviousHighlightedReference");
			AddControlShiftCommand(Key.Up, "ExecuteNextHighlightedReference");

			AddCommand(Key.Back, "ExecuteBackspace");
			AddCommand(Key.Delete, "ExecuteDelete");
			AddControlCommand(Key.Back, "ExecuteWordDeleteToStart");
			AddControlCommand(Key.Delete, "ExecuteWordDeleteToEnd");

			AddCommand(Key.Enter, "ExecuteReturn");
			AddCommand(Key.Tab, "ExecuteTab");
			AddShiftCommand(Key.Tab, "ExecuteBackTab");

			AddControlCommand(Key.V, "ExecutePaste");

			AddIntCommand(ModifierKeys.Control, Key.Z, "ExecuteUndo", 1);
			AddIntCommand(ModifierKeys.Control, Key.Y, "ExecuteRedo", 1);
			AddIntCommand(ModifierKeys.Control | ModifierKeys.Shift, Key.Z, "ExecuteRedo", 1);

			// TODO: Export IDocumentNavigationService to allow F12
			// TODO: Invoke peek from IntellisenseCommandFilter
			// TODO: ExecuteCommentBlock, ExecuteFormatSelection, ExecuteFormatDocument, ExecuteInsertSnippet, ExecuteInsertComment, ExecuteSurroundWith
		}
		#endregion
	}

	[Export(typeof(IChainedKeyProcessorProvider))]
	[TextViewRole(PredefinedTextViewRoles.Interactive)]
	[ContentType("Roslyn Languages")]
	[Order(Before = "Standard KeyProcessor")]
	[Name("Roslyn KeyProcessor")]
	sealed class RoslynKeyProcessorProvider : IChainedKeyProcessorProvider {
		// I need to use the MEF container directly to import a private Roslyn interface.
		readonly IComponentModel componentModel;

		[ImportingConstructor]
		public RoslynKeyProcessorProvider(SVsServiceProvider sp) {
			var mySP = (VsServiceProvider)sp;
			componentModel = mySP.ComponentModel;

			// This is necessary for icons in IntelliSense
			// The initialization is re-entrant, which can
			// make MEF try to create this class again and
			// throw a re-entrancy exception from Lazy<T>.
			// I therefore initialize this asynchronously.
			Dispatcher.CurrentDispatcher.BeginInvoke(new Action(delegate {
				var imageService = new VsImageService(sp);
				imageService.InitializeLibrary();

				// This is necessary for preview icons in CTP3, which doesn't
				// have a singleton CrispImage.DefaultImageLibrary.  The GUID
				// is from SVsImageService, which has a private PIA.
				mySP.AddService(new Guid("ACC9EB93-CAD8-41DE-80DA-BD35CC5112AE"), imageService);
			}));
		}

		public ChainedKeyProcessor GetProcessor(IWpfTextView wpfTextView) {
			return new RoslynKeyProcessor(wpfTextView, componentModel);
		}
	}
}
