using System;
using Microsoft.TeamFoundation.Diff;

namespace AlekseyNagovitsyn.TfsPendingChangesMargin
{
    /// <summary>
    /// Represents information about a specific difference between two sequences.
    /// </summary>
    internal class DiffChange : IDiffChange
    {
        /// <summary>
        /// Creates a <see cref="DiffChange"/> based on the <see cref="IDiffChange"/>.
        /// </summary>
        /// <param name="diffChange">Information about a specific difference between two sequences.</param>
        public DiffChange(IDiffChange diffChange)
        {
            ChangeType = diffChange.ChangeType;

            OriginalStart = diffChange.OriginalStart;
            OriginalEnd = diffChange.OriginalEnd - 1; // exclusive bound to inclusive.
            OriginalLength = diffChange.OriginalLength;

            ModifiedStart = diffChange.ModifiedStart;
            ModifiedEnd = diffChange.ModifiedEnd - 1; // exclusive bound to inclusive.
            ModifiedLength = diffChange.ModifiedLength;
        }

        /// <summary>
        /// The type of difference.
        /// </summary>
        public DiffChangeType ChangeType { get; set; }

        /// <summary>
        /// The inclusive position of the first element in the original sequence which this change affects.
        /// </summary>
        public int OriginalStart { get; set; }

        /// <summary>
        /// The number of elements from the original sequence which were affected (deleted).
        /// </summary>
        public int OriginalLength { get; set; }

        /// <summary>
        /// The inclusive position of the last element in the original sequence which this change affects.
        /// </summary>
        public int OriginalEnd { get; set; }

        /// <summary>
        /// The inclusive position of the first element in the modified sequence which this change affects.
        /// </summary>
        public int ModifiedStart { get; set; }

        /// <summary>
        /// The number of elements from the modified sequence which were affected (added).
        /// </summary>
        public int ModifiedLength { get; set; }

        /// <summary>
        /// The inclusive position of the last element in the modified sequence which this change affects.
        /// </summary>
        public int ModifiedEnd { get; set; }

        /// <summary>
        /// This methods combines two <see cref="IDiffChange"/> objects into one.
        /// </summary>
        /// <param name="diffChange">The diff change to add.</param>
        /// <returns>An <see cref="IDiffChange"/> that represents <c>this</c> + <see cref="diffChange"/>.</returns>
        public IDiffChange Add(IDiffChange diffChange)
        {
            throw new NotImplementedException();
        }
    }
}