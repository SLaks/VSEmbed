using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.Utilities;
using VSEmbed.Editor;

namespace VSEmbed.Exports {
	// Loosely based on WebMatrix's DefaultKeyProcessor
	///<summary>A base class for a KeyProcessor that handles shortcut keys.</summary>
	///<remarks>Derived classes should call Add*Command() in their constructors to build the key map.</remarks>
	public abstract class BaseShortcutKeyProcessor : ChainedKeyProcessor {
		readonly Dictionary<Tuple<ModifierKeys, Key>, Func<bool>> shortcuts = new Dictionary<Tuple<ModifierKeys, Key>, Func<bool>>();

		///<summary>Wraps a void-returning method in an <see cref="Action{Boolean}"/> for the AddCommand methods.</summary>
		protected static Func<bool> WithTrue(Action method) { return () => { method(); return true; }; }
		///<summary>Binds a key which will be invoked with a boolean indicating the state of the Shift key.  This should be used for cursor movement commands.</summary>
		protected void AddExtendableCommand(Key key, Action<bool> method) {
			AddCommand(key, WithTrue(() => method(false)));
			AddShiftCommand(key, WithTrue(() => method(true)));
		}
		///<summary>Binds a key with the Control key, which will be invoked with a boolean indicating the state of the Shift key.  This should be used for cursor movement commands.</summary>
		protected void AddControlExtendableCommand(Key key, Action<bool> method) {
			AddControlCommand(key, WithTrue(() => method(false)));
			AddControlShiftCommand(key, WithTrue(() => method(true)));
		}
		///<summary>Binds a single key.</summary>
		protected void AddCommand(Key key, Func<bool> method) {
			AddCommand(ModifierKeys.None, key, method);
		}
		///<summary>Binds a key with Shift held down.</summary>
		protected void AddShiftCommand(Key key, Func<bool> method) {
			AddCommand(ModifierKeys.Shift, key, method);
		}
		///<summary>Binds a key with Control held down.</summary>
		protected void AddControlCommand(Key key, Func<bool> method) {
			AddCommand(ModifierKeys.Control, key, method);
		}
		///<summary>Binds a key with Control+Shift held down.</summary>
		protected void AddControlShiftCommand(Key key, Func<bool> method) {
			AddCommand(ModifierKeys.Control | ModifierKeys.Shift, key, method);
		}
		///<summary>Binds a key with Alt+Shift held down.</summary>
		protected void AddAltShiftCommand(Key key, Func<bool> method) {
			AddCommand(ModifierKeys.Alt | ModifierKeys.Shift, key, method);
		}
		///<summary>Binds a key with the specified modifiers held down.</summary>
		protected void AddCommand(ModifierKeys modifiers, Key key, Func<bool> method) {
			shortcuts.Add(new Tuple<ModifierKeys, Key>(modifiers, key), method);
		}

		///<summary>Handles the KeyDown event if the pressed key has been bound.</summary>
		public override void KeyDown(KeyEventArgs args, ITextBuffer targetBuffer, Action next) {
			Func<bool> method;
			if (!shortcuts.TryGetValue(Tuple.Create(args.KeyboardDevice.Modifiers, args.Key), out method)
			 || !method())
				next();
		}
	}

	class StandardKeyProcessor : BaseShortcutKeyProcessor {
		readonly IWpfTextView textView;
		readonly IEditorOperations editorOperations;
		readonly ITextUndoHistory undoHistory;

		#region Shortcuts
		void AddShortcuts() {
			// TODO: Add MoveSelectedLinesUp/Down (in IEditorOperations2, which doesn't exist in v11)
			AddExtendableCommand(Key.Right, editorOperations.MoveToNextCharacter);
			AddExtendableCommand(Key.Left, editorOperations.MoveToPreviousCharacter);
			AddControlExtendableCommand(Key.Right, editorOperations.MoveToNextWord);
			AddControlExtendableCommand(Key.Left, editorOperations.MoveToPreviousWord);

			AddExtendableCommand(Key.Up, editorOperations.MoveLineUp);
			AddExtendableCommand(Key.Down, editorOperations.MoveLineDown);
			AddControlCommand(Key.Up, WithTrue(editorOperations.ScrollUpAndMoveCaretIfNecessary));
			AddControlCommand(Key.Down, WithTrue(editorOperations.ScrollDownAndMoveCaretIfNecessary));

			AddExtendableCommand(Key.Home, editorOperations.MoveToHome);
			AddExtendableCommand(Key.End, editorOperations.MoveToEndOfLine);
			AddControlExtendableCommand(Key.Home, editorOperations.MoveToStartOfDocument);
			AddControlExtendableCommand(Key.End, editorOperations.MoveToEndOfDocument);

			AddExtendableCommand(Key.PageUp, editorOperations.PageUp);
			AddExtendableCommand(Key.PageDown, editorOperations.PageDown);
			AddControlExtendableCommand(Key.PageUp, editorOperations.MoveToTopOfView);
			AddControlExtendableCommand(Key.PageDown, editorOperations.MoveToBottomOfView);

			AddControlCommand(Key.U, editorOperations.MakeLowercase);
			AddControlShiftCommand(Key.U, editorOperations.MakeUppercase);

			AddCommand(Key.Back, editorOperations.Backspace);
			AddCommand(Key.Delete, editorOperations.Delete);
			AddControlCommand(Key.Back, editorOperations.DeleteWordToLeft);
			AddControlCommand(Key.Delete, editorOperations.DeleteWordToRight);

			AddCommand(Key.Escape, WithTrue(editorOperations.ResetSelection));
			AddControlCommand(Key.A, WithTrue(editorOperations.SelectAll));
			AddControlCommand(Key.OemPlus, WithTrue(editorOperations.ZoomIn));
			AddControlCommand(Key.OemMinus, WithTrue(editorOperations.ZoomOut));
			AddControlCommand(Key.D0, WithTrue(() => editorOperations.ZoomTo(100)));

			AddCommand(Key.Tab, editorOperations.Indent);
			AddShiftCommand(Key.Tab, editorOperations.Unindent);
			AddCommand(Key.Return, editorOperations.InsertNewLine);
			AddControlCommand(Key.Return, editorOperations.OpenLineAbove);
			AddControlShiftCommand(Key.Return, editorOperations.OpenLineBelow);

			AddControlCommand(Key.X, editorOperations.CutSelection);
			AddControlCommand(Key.C, editorOperations.CopySelection);
			AddControlCommand(Key.V, editorOperations.Paste);
			AddControlShiftCommand(Key.L, editorOperations.DeleteFullLine);
			AddControlCommand(Key.T, editorOperations.TransposeCharacter);
			AddControlShiftCommand(Key.T, editorOperations.TransposeWord);

			// There is currently no way for me to tell whether 
			// a transaction is available without reflection.
			Func<bool> undo = () => {
				try {
					undoHistory.Undo(1);
					return true;
				} catch { return false; }
			};
			Func<bool> redo = () => {
				try {
					undoHistory.Redo(1);
					return true;
				} catch { return false; }
			};
			AddControlCommand(Key.Z, undo);
			AddControlCommand(Key.Y, redo);
			AddControlShiftCommand(Key.Z, redo);
		}
		#endregion

		public StandardKeyProcessor(IWpfTextView textView, IEditorOperations editorOperations, ITextUndoHistory undoHistory) {
			this.textView = textView;
			this.editorOperations = editorOperations;
			this.undoHistory = undoHistory;
			AddShortcuts();
		}

		// TODO: Handle Alt+Arrows to enter box selection


		// Copied directly from WebMatrix's DefaultKeyProcessor
		public override void TextInput(TextCompositionEventArgs args, ITextBuffer targetBuffer, Action next) {
			if (args.Text.Length > 0 && editorOperations.InsertText(args.Text))
				textView.Caret.EnsureVisible();
			else
				next();
		}

		public override void TextInputStart(TextCompositionEventArgs args, ITextBuffer targetBuffer, Action next) {
			if (!(args.TextComposition is ImeTextComposition) || !HandleProvisionalImeInput(args))
				next();
		}
		public override void TextInputUpdate(TextCompositionEventArgs args, ITextBuffer targetBuffer, Action next) {
			if (!(args.TextComposition is ImeTextComposition) || !HandleProvisionalImeInput(args))
				next();
		}
		bool HandleProvisionalImeInput(TextCompositionEventArgs args) {
			if (args.Text.Length == 0 || !editorOperations.InsertProvisionalText(args.Text))
				return false;
			textView.Caret.EnsureVisible();
			return true;
		}
	}


	[Export(typeof(IChainedKeyProcessorProvider))]
	[ContentType("text")]
	[TextViewRole(PredefinedTextViewRoles.Interactive)]
	[Name("Standard KeyProcessor")]
	sealed class StandardKeyProcessorProvider : IChainedKeyProcessorProvider {
		[Import]
		public IEditorOperationsFactoryService EditorOperationsFactory { get; set; }
		[Import]
		public ITextUndoHistoryRegistry UndoHistoryRegistry { get; set; }

		//I'm limiting us to a single keyprocessor and therefore a single wpfTextView
		private StandardKeyProcessor keyProcessor = null;

		public ChainedKeyProcessor GetProcessor(IWpfTextView wpfTextView)
		{
			if (keyProcessor == null)
			{
				keyProcessor = new StandardKeyProcessor(wpfTextView, EditorOperationsFactory.GetEditorOperations(wpfTextView), UndoHistoryRegistry.GetHistory(wpfTextView.TextBuffer));
			}

			return keyProcessor;
		}
	}
}
