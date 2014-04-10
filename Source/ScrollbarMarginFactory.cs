using System.ComponentModel.Composition;

using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
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
    internal sealed class ScrollbarMarginFactory : IWpfTextViewMarginProvider
    {
#pragma warning disable 0649 // Suppress warnings "Field XYZ is never assigned to, and will always have its default value null".

        /// <summary>
        /// Service that creates, loads, and disposes text documents.
        /// </summary>
        [Import]
        internal ITextDocumentFactoryService TextDocumentFactoryService;

        /// <summary>
        /// Visual Studio service provider.
        /// </summary>
        [Import]
        internal SVsServiceProvider VsServiceProvider;

        /// <summary>
        /// Service that provides the <see cref="IEditorFormatMap"/>.
        /// </summary>
        [Import]
        internal IEditorFormatMapService FormatMapService;

#pragma warning restore 0649

        /// <summary>
        /// Creates an <see cref="IWpfTextViewMargin"/> for the given <see cref="IWpfTextViewHost"/>.
        /// </summary>
        /// <param name="textViewHost">The <see cref="IWpfTextViewHost"/> for which to create the <see cref="IWpfTextViewMargin"/>.</param>
        /// <param name="marginContainer">The container that will contain the newly-created margin.</param>
        /// <returns>The <see cref="ScrollbarMargin"/> instance.</returns>
        public IWpfTextViewMargin CreateMargin(IWpfTextViewHost textViewHost, IWpfTextViewMargin marginContainer)
        {
            var marginCore = textViewHost.TextView.Properties.GetOrCreateSingletonProperty(() => new MarginCore(textViewHost.TextView, TextDocumentFactoryService, VsServiceProvider, FormatMapService));
            return new ScrollbarMargin(textViewHost.TextView, marginContainer, marginCore);
        }
    }
}