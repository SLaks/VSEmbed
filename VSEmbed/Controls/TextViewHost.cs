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

namespace VSEmbed.Controls {
	public class TextViewHost : ContentPresenter {
		public IWpfTextView TextView { get; private set; }
		public TextViewHost() {
			if (VsServiceProvider.Instance.ComponentModel == null) {
				if (VsLoader.IsDesignMode)
					return;
				throw new InvalidOperationException("To use TextViewHost, you must first install a MEF container into the ServiceProvider by calling VsMefContainerBuilder.Initialize().");
			}
			var bufferFactory = VsServiceProvider.Instance.ComponentModel.GetService<ITextBufferFactoryService>();
			var editorFactory = VsServiceProvider.Instance.ComponentModel.GetService<ITextEditorFactoryService>();
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
				var contentType = VsServiceProvider.Instance.ComponentModel
					.GetService<IContentTypeRegistryService>()
					.GetContentType(value);
				TextView.TextBuffer.ChangeContentType(contentType, null);
			}
		}
	}
}

