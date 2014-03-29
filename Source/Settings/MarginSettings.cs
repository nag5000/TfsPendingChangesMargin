using System.Windows.Media;

using Microsoft.VisualStudio.Text.Classification;

namespace AlekseyNagovitsyn.TfsPendingChangesMargin.Settings
{
    /// <summary>
    /// The margin settings.
    /// </summary>
    internal sealed class MarginSettings
    {
        /// <summary>
        /// Map that contains format definitions for Visual Studio editor.
        /// </summary>
        private readonly IEditorFormatMap _formatMap;

        /// <summary>
        /// Brush for <see cref="LineDiffType.Added"/> line margin.
        /// </summary>
        public Brush AddedLineMarginBrush { get; set; }

        /// <summary>
        /// Brush for <see cref="LineDiffType.Modified"/> line margin.
        /// </summary>
        public Brush ModifiedLineMarginBrush { get; set; }

        /// <summary>
        /// Brush for <see cref="LineDiffType.Removed"/> line margin.
        /// </summary>
        public Brush RemovedLineMarginBrush { get; set; }

        /// <summary>
        /// Initializes a new instance of margin settings.
        /// </summary>
        /// <param name="formatMap">Map that contains format definitions for Visual Studio editor.</param>
        public MarginSettings(IEditorFormatMap formatMap)
        {
            _formatMap = formatMap;
            Refresh();
        }
        
        /// <summary>
        /// Refresh settings.
        /// </summary>
        public void Refresh()
        {
            AddedLineMarginBrush = _formatMap.GetForeground<AddedLineMarginFormatDefinition>();
            ModifiedLineMarginBrush = _formatMap.GetForeground<ModifiedLineMarginFormatDefinition>();
            RemovedLineMarginBrush = _formatMap.GetForeground<RemovedLineMarginFormatDefinition>();
        }
    }
}