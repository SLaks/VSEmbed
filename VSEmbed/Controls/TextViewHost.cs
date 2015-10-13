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
	///<summary>A WPF control that embeds a Visual Studio editor.</summary>
	public class TextViewHost : ContentPresenter {
		///<summary>Gets the <see cref="IWpfTextView"/> displayed by the control.</summary>
		public IWpfTextView TextView { get; private set; }
		///<summary>Creates a <see cref="TextViewHost"/>.  <see cref="VsMefContainerBuilder"/> must be set up before creating this.</summary>
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
			TextView.Options.SetOptionValue("TextViewHost/SuggestionMargin", true);
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
		///<summary>Gets all content types registered by the editor.</summary>
		public static IReadOnlyList<string> AvailableContentTypes { get { return availableContentTypes.Value; } }


		///<summary>Gets or sets the ContentType of the embedded <see cref="ITextBuffer"/>.</summary>
		public string ContentType {
			get { return (string)GetValue(ContentTypeProperty); }
			set { SetValue(ContentTypeProperty, value); }
		}

		///<summary>Identifies the <see cref="ContentType"/> dependency property.</summary>
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

		///<summary>Gets or sets the content of the embedded <see cref="ITextBuffer"/>.</summary>
		public string Text {
			get { return (string)GetValue(TextProperty); }
			set { SetValue(TextProperty, value); }
		}

		///<summary>Identifies the <see cref="Text"/> dependency property.</summary>
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

