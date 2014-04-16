using System;
using System.Collections;
using System.Collections.Generic;

using Microsoft.TeamFoundation.Diff;
using Microsoft.VisualStudio.Text;

namespace AlekseyNagovitsyn.TfsPendingChangesMargin
{
    /// <summary>
    /// Collection of differences of the text document in comparison with another version.
    /// </summary>
    internal class DiffLinesCollection : IEnumerable<DiffLineEntry>
    {
        /// <summary>
        /// The differences grouped by type.
        /// </summary>
        private readonly Dictionary<DiffChangeType, List<ITextSnapshotLine>> _dict;

        /// <summary>
        /// Initializes a new instance of the <see cref="DiffLinesCollection"/> class.
        /// </summary>
        public DiffLinesCollection()
        {
            _dict = new Dictionary<DiffChangeType, List<ITextSnapshotLine>>();
            foreach (DiffChangeType diffChangeType in Enum.GetValues(typeof(DiffChangeType)))
                _dict.Add(diffChangeType, new List<ITextSnapshotLine>());
        }

        /// <summary>
        /// Gets the differences of a specified type.
        /// </summary>
        /// <param name="diffChangeType">Type of differences.</param>
        /// <returns>Collection of differences with the specified type.</returns>
        public List<ITextSnapshotLine> this[DiffChangeType diffChangeType]
        {
            get { return _dict[diffChangeType]; }
        }

        /// <summary>
        /// Removes all differences from the collection.
        /// </summary>
        public void Clear()
        {
            foreach (List<ITextSnapshotLine> linesList in _dict.Values)
                linesList.Clear();
        }

        /// <summary>
        /// Returns an enumerator that iterates through the collection.
        /// </summary>
        /// <returns>A <see cref="T:System.Collections.Generic.IEnumerator`1"/> that can be used to iterate through the collection.</returns>
        public IEnumerator<DiffLineEntry> GetEnumerator()
        {
            foreach (DiffChangeType diffChangeType in _dict.Keys)
            {
                List<ITextSnapshotLine> linesList = _dict[diffChangeType];
                foreach (ITextSnapshotLine line in linesList)
                {
                    yield return new DiffLineEntry(diffChangeType, line);
                }
            }
        }

        /// <summary>
        /// Returns an enumerator that iterates through a collection.
        /// </summary>
        /// <returns>An <see cref="T:System.Collections.IEnumerator"/> object that can be used to iterate through the collection.</returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
