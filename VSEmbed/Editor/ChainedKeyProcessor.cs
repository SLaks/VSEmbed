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
		///<summary>Handles the KeyDown event, or passes through to the next processor in the chain.</summary>
		public virtual void KeyDown(KeyEventArgs args, ITextBuffer targetBuffer, Action next) {
			next();
		}
		
		///<summary>Handles the TextInputStart event, or passes through to the next processor in the chain.</summary>
		public virtual void TextInputStart(TextCompositionEventArgs args, ITextBuffer targetBuffer, Action next) {
			next();
		}

		///<summary>Handles the TextInput event, or passes through to the next processor in the chain.</summary>
		public virtual void TextInput(TextCompositionEventArgs args, ITextBuffer targetBuffer, Action next) {
			next();
		}
		///<summary>Handles the TextInputUpdate event, or passes through to the next processor in the chain.</summary>
		public virtual void TextInputUpdate(TextCompositionEventArgs args, ITextBuffer targetBuffer, Action next) {
			next();
		}
	}

	///<summary>A MEF-imported service that provides <see cref="ChainedKeyProcessor"/>s for specific content types.</summary>
	public interface IChainedKeyProcessorProvider {
		///<summary>Gets a <see cref="ChainedKeyProcessor"/> for the specified TextView.</summary>
		ChainedKeyProcessor GetProcessor(IWpfTextView wpfTextView);
	}
}