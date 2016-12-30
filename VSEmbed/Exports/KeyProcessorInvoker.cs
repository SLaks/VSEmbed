using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Utilities;
using Microsoft.VisualStudio.Utilities;
using System.Windows.Input;
using Microsoft.VisualStudio.Text;
using VSEmbed.Editor;

namespace VSEmbed.Exports {
	using ChainedKeyProcessorMetadata = Lazy<IChainedKeyProcessorProvider, IOrderableContentTypeAndTextViewRoleMetadata>;
	sealed class KeyProcessorInvoker : KeyProcessor {
		readonly IWpfTextView wpfTextView;
		readonly IList<ChainedKeyProcessorMetadata> keyProcessorProviders;

		public KeyProcessorInvoker(IWpfTextView wpfTextView, IList<ChainedKeyProcessorMetadata> keyProcessorProviders) {
			this.wpfTextView = wpfTextView;
			this.keyProcessorProviders = keyProcessorProviders;
		}

		bool InvokeChain(Action<ChainedKeyProcessor, ITextBuffer, Action> invoker) {
			bool handled = true;
			int index = -1;
			Action next = null;
			next = delegate {
				while (++index < keyProcessorProviders.Count) {
					var targetPoint = wpfTextView.GetCaretPoint(s =>
						keyProcessorProviders[index].Metadata.ContentTypes.Any(s.ContentType.IsOfType)
					);
					// If the cursor is not in the right buffer, skip this processor.
					if (targetPoint == null)
						continue;
					invoker(
						keyProcessorProviders[index].Value.GetProcessor(wpfTextView),
						targetPoint.Value.Snapshot.TextBuffer,
						next
					);
					// After invoking the method, stop.  If it calls next(), we'll continue the chain.
					return;
				}
				// If we got here, we reached the end of the chain without any handler consuming the event.
				handled = false;
			};
			next();
			return handled;
		}

		#region Method Forwarding
		public override void KeyDown(KeyEventArgs args) {
			args.Handled = InvokeChain((processor, buffer, next) => processor.KeyDown(args, buffer, next));
		}
		public override void TextInput(TextCompositionEventArgs args) {
			args.Handled = InvokeChain((processor, buffer, next) => processor.TextInput(args, buffer, next));
		}
		public override void TextInputStart(TextCompositionEventArgs args) {
			args.Handled = InvokeChain((processor, buffer, next) => processor.TextInputStart(args, buffer, next));
		}
		public override void TextInputUpdate(TextCompositionEventArgs args) {
			args.Handled = InvokeChain((processor, buffer, next) => processor.TextInputUpdate(args, buffer, next));
		}
		#endregion
	}
	[Export(typeof(IKeyProcessorProvider))]
	[TextViewRole(PredefinedTextViewRoles.Interactive)]
	[ContentType("text")]
	[Name("KeyProcessorInvokerProvider")]
	sealed class KeyProcessorInvokerProvider : IKeyProcessorProvider, IPartImportsSatisfiedNotification {

		[ImportMany]
		public IEnumerable<ChainedKeyProcessorMetadata> KeyProcessorProviders { get; set; }

		public IList<ChainedKeyProcessorMetadata> OrderedKeyProcessorProviders { get; private set; }

		public void OnImportsSatisfied() {
			OrderedKeyProcessorProviders = Orderer.Order(KeyProcessorProviders);
		}

		public KeyProcessor GetAssociatedProcessor(IWpfTextView wpfTextView) {
			return new KeyProcessorInvoker(
				wpfTextView,
				OrderedKeyProcessorProviders.Where(p => p.Metadata.TextViewRoles.Any(wpfTextView.Roles.Contains)).ToList()
			);
		}
	}
}
