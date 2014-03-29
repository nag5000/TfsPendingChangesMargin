using System.ComponentModel.Composition;
using System.Windows.Media;

using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Utilities;

namespace AlekseyNagovitsyn.TfsPendingChangesMargin.Settings
{
    /// <summary>
    /// The editor format definition for <see cref="LineDiffType.Added"/> line margin.
    /// </summary>
    [ClassificationType(ClassificationTypeNames = DefinitionKey)]
    [Name(DefinitionKey)]
    [DisplayName(DefinitionTitle)]
    [Order(Before = Priority.Default)]
    [TextViewRole(PredefinedTextViewRoles.Document)]
    [UserVisible(true)]
    [Export(typeof(EditorFormatDefinition))]
    [ContentType("text")]
    internal sealed class AddedLineMarginFormatDefinition : EditorFormatDefinition
    {
        /// <summary>
        /// The definition has been stored in <see cref="IEditorFormatMap"/> by this key.
        /// </summary>
        private const string DefinitionKey = EditorMargin.MarginName + "_AddedLineMargin";

        /// <summary>
        /// Caption of the definition in "Tools -> Environment -> Fonts and Colors" dialog.
        /// </summary>
        private const string DefinitionTitle = EditorMargin.MarginName + " Added line margin";

        /// <summary>
        /// Initializes a new instance of the <see cref="AddedLineMarginFormatDefinition"/> class with default properties.
        /// </summary>
        public AddedLineMarginFormatDefinition()
        {
            DisplayName = DefinitionTitle;
            ForegroundColor = Colors.LawnGreen;
            BackgroundCustomizable = false;
        }
    }
}