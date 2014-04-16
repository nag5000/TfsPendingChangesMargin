using System.ComponentModel.Composition;

using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Editor;

namespace AlekseyNagovitsyn.TfsPendingChangesMargin
{
    /// <summary>
    /// Export a <see cref="IWpfTextViewMarginProvider"/>, which returns an instance of the <see cref="MarginCore"/>-based margin for a <see cref="IWpfTextViewHost"/>.
    /// </summary>
    internal abstract class MarginCoreFactory : IWpfTextViewMarginProvider
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

        /// <summary>
        /// Factory that creates or reuses an <see cref="IScrollMap"/> for an <see cref="ITextView"/>.
        /// </summary>
        [Import]
        internal IScrollMapFactoryService ScrollMapFactoryService;

#pragma warning restore 0649

        /// <summary>
        /// Creates an <see cref="IWpfTextViewMargin"/> for the given <see cref="IWpfTextViewHost"/>.
        /// </summary>
        /// <param name="textViewHost">The <see cref="IWpfTextViewHost"/> for which to create the <see cref="IWpfTextViewMargin"/>.</param>
        /// <param name="marginContainer">The container that will contain the newly-created margin.</param>
        /// <returns>The <see cref="IWpfTextViewMargin"/> instance.</returns>
        public abstract IWpfTextViewMargin CreateMargin(IWpfTextViewHost textViewHost, IWpfTextViewMargin marginContainer);

        /// <summary>
        /// Gets <see cref="MarginCore"/> for the specified <see cref="IWpfTextViewHost"/>.
        /// </summary>
        /// <param name="textViewHost">The <see cref="IWpfTextViewHost"/> for which to create or get the <see cref="MarginCore"/>.</param>
        /// <returns>The <see cref="MarginCore"/> instance.</returns>
        protected MarginCore GetMarginCore(IWpfTextViewHost textViewHost)
        {
            var marginCore = textViewHost.TextView.Properties.GetOrCreateSingletonProperty(() => new MarginCore(textViewHost.TextView, TextDocumentFactoryService, VsServiceProvider, FormatMapService, ScrollMapFactoryService));
            return marginCore;
        }
    }
}