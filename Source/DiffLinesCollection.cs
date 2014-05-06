using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

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
        /// Collection which linking the changed line with change information.
        /// </summary>
        private readonly Dictionary<ITextSnapshotLine, IDiffChange> _diffDict;

        /// <summary>
        /// Initializes a new instance of the <see cref="DiffLinesCollection"/> class.
        /// </summary>
        public DiffLinesCollection()
        {
            _dict = new Dictionary<DiffChangeType, List<ITextSnapshotLine>>();
            foreach (DiffChangeType diffChangeType in Enum.GetValues(typeof(DiffChangeType)))
                _dict.Add(diffChangeType, new List<ITextSnapshotLine>());

            _diffDict = new Dictionary<ITextSnapshotLine, IDiffChange>();
        }

        /// <summary>
        /// Gets the differences of a specified type.
        /// </summary>
        /// <param name="diffChangeType">Type of differences.</param>
        /// <returns>Collection of differences with the specified type.</returns>
        public IReadOnlyList<ITextSnapshotLine> this[DiffChangeType diffChangeType]
        {
            get { return _dict[diffChangeType]; }
        }

        /// <summary>
        /// Gets information about the difference of line in the collection.
        /// </summary>
        /// <param name="line">A line of text from an <see cref="ITextSnapshot"/> in the collection.</param>
        /// <returns>Information about the difference of line.</returns>
        public IDiffChange this[ITextSnapshotLine line]
        {
            get { return _diffDict[line]; }
        }

        /// <summary>
        /// Add the changed line in the collection.
        /// </summary>
        /// <param name="line">A line of text from an <see cref="ITextSnapshot"/>.</param>
        /// <param name="diffChangeInfo">Information about the difference of line.</param>
        public void Add(ITextSnapshotLine line, IDiffChange diffChangeInfo)
        {
            ITextSnapshotLine existsLine = _diffDict.Keys.FirstOrDefault(x => x.LineNumber == line.LineNumber);
            if (existsLine != null)
            {
                IDiffChange oldDiffChangeInfo = _diffDict[existsLine];
                diffChangeInfo = oldDiffChangeInfo.Add(diffChangeInfo);

                _dict[oldDiffChangeInfo.ChangeType].Remove(existsLine);
                _diffDict.Remove(existsLine);
            }

            _dict[diffChangeInfo.ChangeType].Add(line);
            _diffDict[line] = diffChangeInfo;
        }

        /// <summary>
        /// Removes all differences from the collection.
        /// </summary>
        public void Clear()
        {
            foreach (List<ITextSnapshotLine> linesList in _dict.Values)
                linesList.Clear();

            _diffDict.Clear();
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
