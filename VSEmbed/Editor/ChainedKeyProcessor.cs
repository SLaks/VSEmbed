using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;

namespace VSEmbed.Editor {
	///<summary>A base class which handles keystrokes, or passes them on to the next KeyProcessor in the chain.</summary>
	public abstract class ChainedKeyProcessor {
		public virtual void PreviewKeyDown(KeyEventArgs args, ITextBuffer targetBuffer, Action next) {
			next();
		}
		public virtual void KeyDown(KeyEventArgs args, ITextBuffer targetBuffer, Action next) {
			next();
		}
		public virtual void PreviewKeyUp(KeyEventArgs args, ITextBuffer targetBuffer, Action next) {
			next();
		}
		public virtual void KeyUp(KeyEventArgs args, ITextBuffer targetBuffer, Action next) {
			next();
		}
		public virtual void PreviewTextInputStart(TextCompositionEventArgs args, ITextBuffer targetBuffer, Action next) {
			next();
		}
		public virtual void TextInputStart(TextCompositionEventArgs args, ITextBuffer targetBuffer, Action next) {
			next();
		}
		public virtual void PreviewTextInput(TextCompositionEventArgs args, ITextBuffer targetBuffer, Action next) {
			next();
		}
		public virtual void TextInput(TextCompositionEventArgs args, ITextBuffer targetBuffer, Action next) {
			next();
		}
		public virtual void PreviewTextInputUpdate(TextCompositionEventArgs args, ITextBuffer targetBuffer, Action next) {
			next();
		}
		public virtual void TextInputUpdate(TextCompositionEventArgs args, ITextBuffer targetBuffer, Action next) {
			next();
		}
	}

	///<summary>A MEF-imported service that provides <see cref="ChainedKeyProcessor"/>s for specific content types.</summary>
	public interface IChainedKeyProcessorProvider {
		ChainedKeyProcessor GetProcessor(IWpfTextView wpfTextView);
	}
}