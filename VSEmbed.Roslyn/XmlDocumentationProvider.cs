using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Microsoft.CodeAnalysis;

namespace VSEmbed.Roslyn {
	///<summary>
	/// A <see cref="DocumentationProvider"/> that correctly returns XML tags in the content, working around https://roslyn.codeplex.com/workitem/406. 
	///</summary>
	public class XmlDocumentationProvider : DocumentationProvider {
		#pragma warning disable CS1591 // Missing XML comment for publicly visible type or member; consult the Roslyn sources.
		private volatile Dictionary<string, string> _docComments;

		private readonly string _filePath;

		public XmlDocumentationProvider(string filePath) {
			_filePath = filePath;
		}

		public override bool Equals(object obj) {
			var other = obj as XmlDocumentationProvider;
			return other != null && _filePath == other._filePath;
		}

		public override int GetHashCode() {
			return _filePath.GetHashCode();
		}

		protected override string GetDocumentationForSymbol(string documentationMemberID, CultureInfo preferredCulture, CancellationToken cancellationToken = default(CancellationToken)) {
			if (_docComments == null) {
				try {
					using (var reader = XmlReader.Create(_filePath, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit })) {
						_docComments = XDocument.Load(_filePath)
							.Descendants("member")
							.Where(e => e.Attribute("name") != null)
							.ToDictionary(e => e.Attribute("name").Value, e => string.Concat(e.Nodes()));
					}
				} catch (Exception) {
				}
			}

			string docComment;
			return _docComments.TryGetValue(documentationMemberID, out docComment) ? docComment : "";
		}
	}
}
