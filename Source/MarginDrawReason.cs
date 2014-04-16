using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Editor;

namespace AlekseyNagovitsyn.TfsPendingChangesMargin
{
    /// <summary>
    /// Reason of drawing the margin.
    /// </summary>
    internal enum MarginDrawReason
    {
        /// <summary>
        /// Internal reason.
        /// </summary>
        InternalReason,

        /// <summary>
        /// Version control item has been changed.
        /// </summary>
        VersionControlItemChanged,

        /// <summary>
        /// <see cref="IWpfTextView.ZoomLevel"/> has been changed.
        /// </summary>
        TextViewZoomLevelChanged,

        /// <summary>
        /// The document has been loaded from or saved to disk. 
        /// </summary>
        TextDocFileActionOccurred,

        /// <summary>
        /// The document text has been changed.
        /// </summary>
        TextViewTextChanged,

        /// <summary>
        /// The text editor performs a text line layout.
        /// </summary>
        TextViewLayoutChanged,

        /// <summary>
        /// <see cref="IEditorFormatMap"/> has been changed.
        /// </summary>
        EditorFormatMapChanged,

        /// <summary>
        /// The mapping has been changed between a character position and its vertical fraction.
        /// </summary>
        ScrollMapMappingChanged
    }
}