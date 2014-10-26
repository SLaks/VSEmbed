using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using VSThemeBrowser.VisualStudio;

namespace VSThemeBrowser.Controls {
	public class TextViewHost : ContentPresenter {
		public IWpfTextView TextView { get; private set; }
		static Guid TextViewHostProperty = DefGuidList.guidIWpfTextViewHost;
		public TextViewHost() {
			//var bufferFactory = Mef.Container.GetExportedValue<ITextBufferFactoryService>();
			var editorFactory = Mef.Container.GetExportedValue<ITextEditorFactoryService>();
			var adapterFactory = Mef.Container.GetExportedValue<IVsEditorAdaptersFactoryService>();

			var bufferAdapter = adapterFactory.CreateVsTextBufferAdapter(FakeServiceProvider.Instance);
			var viewAdapter = adapterFactory.CreateVsTextViewAdapter(FakeServiceProvider.Instance, editorFactory.AllPredefinedRoles);

			bufferAdapter.InitializeContent("", 0);		// Force InitializeDocumentTextBuffer()

			viewAdapter.SetBuffer((IVsTextLines)bufferAdapter);
			// According to my decompiler, this is the only way to get the TextViewHost that the SimpleTextViewWindow creates.
			object textViewHostObj;
			((IVsUserData)viewAdapter).GetData(ref TextViewHostProperty, out textViewHostObj);

			var textViewHost = ((IWpfTextViewHost)textViewHostObj);
			Content = textViewHost.HostControl;
			TextView = textViewHost.TextView;
		}

		public string Text {
			get { return TextView.TextSnapshot.GetText(); }
			set {
				TextView.TextBuffer.Replace(new Span(0, TextView.TextSnapshot.Length), value);
			}
		}
	}
}
