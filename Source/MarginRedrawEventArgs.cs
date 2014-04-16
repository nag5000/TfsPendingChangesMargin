using System;

namespace AlekseyNagovitsyn.TfsPendingChangesMargin
{
    /// <summary>
    /// Arguments for event that raised when margin needs to be redrawn.
    /// </summary>
    internal class MarginRedrawEventArgs : EventArgs
    {
        /// <summary>
        /// Differences between the current document and the version in TFS.
        /// </summary>
        private readonly DiffLinesCollection _diffLines;

        /// <summary>
        /// The reason of redrawing the margin.
        /// </summary>
        private readonly MarginDrawReason _reason;

        /// <summary>
        /// Gets differences between the current document and the version in TFS.
        /// </summary>
        public DiffLinesCollection DiffLines
        {
            get { return _diffLines; }
        }

        /// <summary>
        /// Gets the reason of redrawing the margin.
        /// </summary>
        public MarginDrawReason Reason
        {
            get { return _reason; }
        }

        /// <summary>
        /// Initialize a new instance of the <see cref="MarginRedrawEventArgs"/> class.
        /// </summary>
        /// <param name="diffLines">Differences between the current document and the version in TFS.</param>
        /// <param name="reason">The reason of redrawing the margin.</param>
        public MarginRedrawEventArgs(DiffLinesCollection diffLines, MarginDrawReason reason)
        {
            _diffLines = diffLines;
            _reason = reason;
        }
    }
}