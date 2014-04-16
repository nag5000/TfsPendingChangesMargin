using System.ComponentModel.Composition;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Utilities;

namespace AlekseyNagovitsyn.TfsPendingChangesMargin
{
    /// <summary>
    /// Export a <see cref="IWpfTextViewMarginProvider"/>, which returns an instance of the margin for the editor scrollbar to use.
    /// </summary>
    [Export(typeof(IWpfTextViewMarginProvider))]
    [Name(ScrollbarMargin.MarginName)]
    [ContentType("text")] // Show this margin for all text-based types.
    [TextViewRole(PredefinedTextViewRoles.PrimaryDocument)]
    [MarginContainer(PredefinedMarginNames.VerticalScrollBar)]
    [Order(After = PredefinedMarginNames.OverviewChangeTracking)]
    [Order(Before = PredefinedMarginNames.OverviewError)]
    [Order(Before = PredefinedMarginNames.OverviewMark)]
    [Order(Before = PredefinedMarginNames.OverviewSourceImage)]
    internal sealed class ScrollbarMarginFactory : MarginCoreFactory
    {
        /// <summary>
        /// Creates an <see cref="IWpfTextViewMargin"/> for the given <see cref="IWpfTextViewHost"/>.
        /// </summary>
        /// <param name="textViewHost">The <see cref="IWpfTextViewHost"/> for which to create the <see cref="IWpfTextViewMargin"/>.</param>
        /// <param name="marginContainer">The container that will contain the newly-created margin.</param>
        /// <returns>The <see cref="ScrollbarMargin"/> instance.</returns>
        public override IWpfTextViewMargin CreateMargin(IWpfTextViewHost textViewHost, IWpfTextViewMargin marginContainer)
        {
            MarginCore marginCore = GetMarginCore(textViewHost);
            return new ScrollbarMargin(textViewHost.TextView, marginContainer, marginCore);
        }
    }
}