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
            OriginalEnd = Math.Max(0, diffChange.OriginalEnd - 1); // exclusive bound to inclusive.
            OriginalLength = diffChange.OriginalLength;

            if (ChangeType == DiffChangeType.Delete)
            {
                ModifiedStart = Math.Max(0, diffChange.ModifiedEnd - 1); // exclusive upper bound.
                ModifiedEnd = Math.Max(0, diffChange.ModifiedStart - 1); // exclusive lower bound.
            }
            else
            {
                ModifiedStart = diffChange.ModifiedStart;
                ModifiedEnd = Math.Max(0, diffChange.ModifiedEnd - 1); // exclusive bound to inclusive.
            }

            ModifiedLength = diffChange.ModifiedLength;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DiffChange"/> class.
        /// </summary>
        private DiffChange()
        {
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
        /// The inclusive (exclusive for <see cref="DiffChangeType.Delete"/>) position 
        /// of the first element in the modified sequence which this change affects.
        /// </summary>
        public int ModifiedStart { get; set; }

        /// <summary>
        /// The number of elements from the modified sequence which were affected (added).
        /// </summary>
        public int ModifiedLength { get; set; }

        /// <summary>
        /// The inclusive (exclusive for <see cref="DiffChangeType.Delete"/>) position 
        /// of the last element in the modified sequence which this change affects.
        /// </summary>
        public int ModifiedEnd { get; set; }

        /// <summary>
        /// This methods combines two <see cref="IDiffChange"/> objects into one.
        /// </summary>
        /// <param name="diffChange">The diff change to add.</param>
        /// <returns>A new instance of <see cref="IDiffChange"/> that represents <c>this</c> + <see cref="diffChange"/>.</returns>
        public IDiffChange Add(IDiffChange diffChange)
        {
            if (diffChange == null)
                return this;

            int originalStart = Math.Min(OriginalStart, diffChange.OriginalStart);
            int originalEnd = Math.Max(OriginalEnd, diffChange.OriginalEnd);
            int modifiedStart = Math.Min(ModifiedStart, diffChange.ModifiedStart);
            int modifiedEnd = Math.Max(ModifiedEnd, diffChange.ModifiedEnd);

            DiffChangeType changeType;
            if (ChangeType == diffChange.ChangeType && OriginalStart - diffChange.OriginalEnd == 0 && ModifiedStart - diffChange.ModifiedEnd == 0) 
                changeType = ChangeType;
            else 
                changeType = DiffChangeType.Change;

            return new DiffChange
            {
                ChangeType = changeType,
                OriginalStart = originalStart,
                OriginalEnd = originalEnd,
                OriginalLength = originalEnd - originalStart,
                ModifiedStart = modifiedStart,
                ModifiedEnd = modifiedEnd,
                ModifiedLength = modifiedEnd - modifiedStart
            };
        }
    }
}