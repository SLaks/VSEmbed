using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.Utilities;
using VSThemeBrowser.VisualStudio;

namespace VSThemeBrowser.Controls {
	public class TextViewHost : ContentPresenter {
		public IWpfTextView TextView { get; private set; }
		public TextViewHost() {
			var bufferFactory = VsMefContainerBuilder.Container.GetExportedValue<ITextBufferFactoryService>();
			var editorFactory = VsMefContainerBuilder.Container.GetExportedValue<ITextEditorFactoryService>();
			TextView = editorFactory.CreateTextView(
				bufferFactory.CreateTextBuffer(),
				editorFactory.AllPredefinedRoles
			);
			Content = editorFactory.CreateTextViewHost(TextView, false).HostControl;
		}

		public string Text {
			get { return TextView.TextSnapshot.GetText(); }
			set {
				TextView.TextBuffer.Replace(new Span(0, TextView.TextSnapshot.Length), value);
			}
		}
		public string ContentType {
			get { return TextView.TextBuffer.ContentType.TypeName; }
			set {
				var contentType = VsMefContainerBuilder.Container.GetExportedValue<IContentTypeRegistryService>().GetContentType(value);
				TextView.TextBuffer.ChangeContentType(contentType, null);
			}
		}
	}
	// Loosely based on WebMatrix's DefaultKeyProcessor
	class SimpleKeyProcessor : KeyProcessor {
		readonly IWpfTextView textView;
		readonly IEditorOperations editorOperations;
		readonly ITextUndoHistory undoHistory;

		readonly Dictionary<Tuple<ModifierKeys, Key>, Func<bool>> shortcuts = new Dictionary<Tuple<ModifierKeys, Key>, Func<bool>>();
		#region Shortcuts
		private void AddShortcuts() {
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

			AddControlCommand(Key.U, editorOperations.MakeUppercase);
			AddControlShiftCommand(Key.U, editorOperations.MakeLowercase);

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

		static Func<bool> WithTrue(Action method) { return () => { method(); return true; }; }
		private void AddExtendableCommand(Key key, Action<bool> method) {
			AddCommand(key, WithTrue(() => method(false)));
			AddShiftCommand(key, WithTrue(() => method(true)));
		}
		private void AddControlExtendableCommand(Key key, Action<bool> method) {
			AddControlCommand(key, WithTrue(() => method(false)));
			AddControlShiftCommand(key, WithTrue(() => method(true)));
		}
		private void AddCommand(Key key, Func<bool> method) {
			AddCommand(ModifierKeys.None, key, method);
		}
		private void AddShiftCommand(Key key, Func<bool> method) {
			AddCommand(ModifierKeys.Shift, key, method);
		}
		private void AddControlCommand(Key key, Func<bool> method) {
			AddCommand(ModifierKeys.Control, key, method);
		}
		private void AddControlShiftCommand(Key key, Func<bool> method) {
			AddCommand(ModifierKeys.Control | ModifierKeys.Shift, key, method);
		}
		private void AddAltShiftCommand(Key key, Func<bool> method) {
			AddCommand(ModifierKeys.Alt | ModifierKeys.Shift, key, method);
		}
		private void AddCommand(ModifierKeys modifiers, Key key, Func<bool> method) {
			shortcuts.Add(new Tuple<ModifierKeys, Key>(modifiers, key), method);
		}
		#endregion


		public SimpleKeyProcessor(IWpfTextView textView, IEditorOperations editorOperations, ITextUndoHistory undoHistory) {
			this.textView = textView;
			this.editorOperations = editorOperations;
			this.undoHistory = undoHistory;
			AddShortcuts();
		}

		public override void KeyDown(KeyEventArgs args) {
			base.KeyDown(args);
			Func<bool> method;
			args.Handled = shortcuts.TryGetValue(Tuple.Create(args.KeyboardDevice.Modifiers, args.Key), out method)
				&& method();
			// TODO: Handle Alt+Arrows to enter box selection
		}

		public override void TextInput(TextCompositionEventArgs args) {
			if (args.Text.Length > 0)
				args.Handled = editorOperations.InsertText(args.Text);
			if (args.Handled)
				textView.Caret.EnsureVisible();
		}


		public override void TextInputStart(TextCompositionEventArgs args) {
			if (args.TextComposition is ImeTextComposition) {
				HandleProvisionalImeInput(args);
			}
		}
		public override void TextInputUpdate(TextCompositionEventArgs args) {
			if (args.TextComposition is ImeTextComposition)
				HandleProvisionalImeInput(args);
			else
				args.Handled = false;
		}
		private void HandleProvisionalImeInput(TextCompositionEventArgs args) {
			if (args.Text.Length == 0)
				return;
			args.Handled = editorOperations.InsertProvisionalText(args.Text);
			if (args.Handled)
				textView.Caret.EnsureVisible();
		}
	}
}

