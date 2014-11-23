using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
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

			TextView.TextBuffer.ContentTypeChanged += (s, e) => ContentType = TextView.TextBuffer.ContentType.DisplayName;
			TextView.TextBuffer.Changed += (s, e) => Text = TextView.TextSnapshot.GetText();
		}

		private static Lazy<IReadOnlyList<string>> availableContentTypes = new Lazy<IReadOnlyList<string>>(() =>
			VsServiceProvider.Instance.ComponentModel
				.GetService<IContentTypeRegistryService>()
				.ContentTypes
				.Select(c => c.DisplayName)
				.ToList()
				.AsReadOnly()
		);
		public static IReadOnlyList<string> AvailableContentTypes { get { return availableContentTypes.Value; } }


		public string ContentType {
			get { return (string)GetValue(ContentTypeProperty); }
			set { SetValue(ContentTypeProperty, value); }
		}

		public static readonly DependencyProperty ContentTypeProperty =
			DependencyProperty.Register("ContentType", typeof(string), typeof(TextViewHost), new PropertyMetadata(ContentType_Changed));
		private static void ContentType_Changed(DependencyObject d, DependencyPropertyChangedEventArgs e) {
			var instance = (TextViewHost)d;
			var contentTypeRegistry = VsServiceProvider.Instance.ComponentModel.GetService<IContentTypeRegistryService>();

			var contentType =
				contentTypeRegistry.GetContentType(e.NewValue as string ?? "text")
			 ?? contentTypeRegistry.GetContentType("text");

			instance.TextView.TextBuffer.ChangeContentType(contentType, null);
		}

		public string Text {
			get { return (string)GetValue(TextProperty); }
			set { SetValue(TextProperty, value); }
		}

		// Using a DependencyProperty as the backing store for Text.  This enables animation, styling, binding, etc...
		public static readonly DependencyProperty TextProperty =
			DependencyProperty.Register("Text", typeof(string), typeof(TextViewHost), new PropertyMetadata(Text_Changed));
		private static void Text_Changed(DependencyObject d, DependencyPropertyChangedEventArgs e) {
			var instance = (TextViewHost)d;
			var value = e.NewValue as string;
			if (value != instance.Text)
				instance.TextView.TextBuffer.Replace(new Span(0, instance.TextView.TextSnapshot.Length), value);
		}
	}
}

