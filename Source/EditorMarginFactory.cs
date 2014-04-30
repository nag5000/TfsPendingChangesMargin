using System.ComponentModel.Composition;

using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.Utilities;

namespace AlekseyNagovitsyn.TfsPendingChangesMargin
{
    /// <summary>
    /// Export a <see cref="IWpfTextViewMarginProvider"/>, which returns an instance of the margin for the editor to use.
    /// </summary>
    [Export(typeof(IWpfTextViewMarginProvider))]
    [Name(EditorMargin.MarginName)]
    [Order(After = PredefinedMarginNames.Spacer)]
    [Order(Before = PredefinedMarginNames.Outlining)]
    [MarginContainer(PredefinedMarginNames.LeftSelection)]
    [ContentType("text")] // Show this margin for all text-based types.
    [TextViewRole(PredefinedTextViewRoles.PrimaryDocument)]
    internal sealed class EditorMarginFactory : MarginCoreFactory
    {
        /// <summary>
        /// Maintains the relationship between text buffers and <see cref="ITextUndoHistory"/> objects.
        /// </summary>
        [Import]
        internal ITextUndoHistoryRegistry UndoHistoryRegistry;

        /// <summary>
        /// Creates an <see cref="IWpfTextViewMargin"/> for the given <see cref="IWpfTextViewHost"/>.
        /// </summary>
        /// <param name="textViewHost">The <see cref="IWpfTextViewHost"/> for which to create the <see cref="IWpfTextViewMargin"/>.</param>
        /// <param name="marginContainer">The container that will contain the newly-created margin.</param>
        /// <returns>The <see cref="EditorMargin"/> instance.</returns>
        public override IWpfTextViewMargin CreateMargin(IWpfTextViewHost textViewHost, IWpfTextViewMargin marginContainer)
        {
            var marginCore = GetMarginCore(textViewHost);
            return new EditorMargin(textViewHost.TextView, UndoHistoryRegistry, marginCore);
        }
    }
}
