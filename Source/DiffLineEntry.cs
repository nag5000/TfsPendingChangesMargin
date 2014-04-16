using Microsoft.TeamFoundation.Diff;
using Microsoft.VisualStudio.Text;

namespace AlekseyNagovitsyn.TfsPendingChangesMargin
{
    /// <summary>
    /// Difference item of the text document in comparison with another version.
    /// </summary>
    internal class DiffLineEntry
    {
        /// <summary>
        /// Difference type.
        /// </summary>
        private readonly DiffChangeType _changeType;

        /// <summary>
        /// Text line which differs.
        /// </summary>
        private readonly ITextSnapshotLine _line;

        /// <summary>
        /// Gets difference type.
        /// </summary>
        public DiffChangeType ChangeType
        {
            get { return _changeType; }
        }

        /// <summary>
        /// Gets text line which differs.
        /// </summary>
        public ITextSnapshotLine Line
        {
            get { return _line; }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DiffLineEntry"/> class.
        /// </summary>
        /// <param name="changeType">Difference type.</param>
        /// <param name="line">Text line which differs.</param>
        public DiffLineEntry(DiffChangeType changeType, ITextSnapshotLine line)
        {
            _changeType = changeType;
            _line = line;
        }
    }
}