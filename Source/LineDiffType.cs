namespace AlekseyNagovitsyn.TfsPendingChangesMargin
{
    /// <summary>
    /// A type of difference of linked lines.
    /// </summary>
    internal enum LineDiffType
    {
        /// <summary>
        /// Line has been added.
        /// </summary>
        Added,

        /// <summary>
        /// Line has been modified.
        /// </summary>
        Modified,

        /// <summary>
        /// Line has been removed.
        /// </summary>
        Removed
    }
}