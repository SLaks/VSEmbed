using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel.Composition;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.Utilities;
using VSEmbed.Editor;
using VSEmbed.Exports;

namespace VSEmbed.Roslyn {
	// This class has nothing to do with Roslyn, but it consumes APIs new to Dev14,
	// so I can't put it elsewhere.  And these features are only provided by Roslyn
	// anyway, so it doesn't matter.  This is based on IntellisenseCommandFilter in
	// Microsoft.VisualStudio.Editor.Implementation
	[Export(typeof(ChainedKeyProcessor))]
	class Dev14KeyProcessor : BaseShortcutKeyProcessor {
		readonly IWpfTextView wpfTextView;
		readonly ILightBulbBroker lightBulbBroker;
		readonly ISuggestedActionCategoryRegistryService suggestedActionCategoryRegistryService;
		readonly ISmartTagBroker smartTagBroker;
		readonly IPeekBroker peekBroker;

		public Dev14KeyProcessor(IWpfTextView wpfTextView, ILightBulbBroker lightBulbBroker, ISuggestedActionCategoryRegistryService suggestedActionCategoryRegistryService, ISmartTagBroker smartTagBroker, IPeekBroker peekBroker) {
			this.wpfTextView = wpfTextView;
			this.lightBulbBroker = lightBulbBroker;
			this.suggestedActionCategoryRegistryService = suggestedActionCategoryRegistryService;
			this.smartTagBroker = smartTagBroker;
			this.peekBroker = peekBroker;

			AddShortcuts();
		}

		void AddShortcuts() {
			AddControlCommand(Key.OemPeriod, TryShowSuggestedActions);

			// Alt+regular keys doesn't trigger this at all
			AddCommand(Key.F12, TryPeek);
		}

		bool TryShowSuggestedActions() {
			if (lightBulbBroker.TryExpandSession(suggestedActionCategoryRegistryService.AllCodeFixes, wpfTextView))
				return true;
			return ExpandSmartTagUnderCaret();
		}
		private bool ExpandSmartTagUnderCaret() {
			var insertionPoint = wpfTextView.Caret.Position.Point.GetInsertionPoint(b => b == wpfTextView.TextBuffer);
			var snapshot = insertionPoint.Value.Snapshot;
			var caretSpan = new SnapshotSpan(insertionPoint.Value, 0);
			ReadOnlyCollection<ISmartTagSession> sessions = smartTagBroker.GetSessions(wpfTextView);
			ISmartTagSession session =
				sessions.FirstOrDefault(s => s.ApplicableToSpan.GetSpan(snapshot).IntersectsWith(caretSpan) && s.Type == SmartTagType.Factoid)
			 ?? sessions.FirstOrDefault(s => s.ApplicableToSpan.GetSpan(snapshot).IntersectsWith(caretSpan) && s.Type == SmartTagType.Ephemeral)
			 ?? sessions.FirstOrDefault(s => s.Type == SmartTagType.Ephemeral);
			// VS will only ignore the caret if there is exactly one ephemeral session, but I'm lazy
			if (session == null)
				return false;
			session.State = SmartTagState.Expanded;
			return true;
		}

		bool TryPeek() {
			return peekBroker.TriggerPeekSession(wpfTextView, PredefinedPeekRelationships.Definitions.Name) != null;
		}
	}
	[Export(typeof(IChainedKeyProcessorProvider))]
	[ContentType("text")]
	[TextViewRole(PredefinedTextViewRoles.Interactive)]
	[Name("Dev14 KeyProcessor")]
	sealed class Dev14KeyProcessorProvider : IChainedKeyProcessorProvider {
		[Import]
		public ILightBulbBroker LightBulbBroker { get; set; }

		[Import]
		public ISmartTagBroker SmartTagBroker { get; set; }

		[Import]
		public ISuggestedActionCategoryRegistryService SuggestedActionCategoryRegistryService { get; set; }

		[Import]
		public IPeekBroker PeekBroker { get; set; }

		//I'm limiting us to a single keyprocessor and therefore a single wpfTextView
		private Dev14KeyProcessor keyProcessor = null;

		public ChainedKeyProcessor GetProcessor(IWpfTextView wpfTextView)
		{
			if (keyProcessor == null)
			{
				return new Dev14KeyProcessor(wpfTextView, LightBulbBroker, SuggestedActionCategoryRegistryService, SmartTagBroker, PeekBroker);
			}

			return keyProcessor;
		}
	}
}
