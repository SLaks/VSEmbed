using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using VSThemeBrowser.VisualStudio;

namespace VSThemeBrowser.Controls {
	public class TextViewHost : ContentPresenter {
		public IWpfTextView TextView { get; private set; }
		public TextViewHost() {
			var bufferFactory = Mef.Container.GetExportedValue<ITextBufferFactoryService>();
			var editorFactory = Mef.Container.GetExportedValue<ITextEditorFactoryService>();
			Content = TextView = editorFactory.CreateTextView(
				bufferFactory.CreateTextBuffer(), 
				editorFactory.AllPredefinedRoles
			);
		}

		public string Text {
			get { return TextView.TextSnapshot.GetText(); }
			set {
				TextView.TextBuffer.Replace(new Span(0, TextView.TextSnapshot.Length), value);
			}
		}
	}
}
