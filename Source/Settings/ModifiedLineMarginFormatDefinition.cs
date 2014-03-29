using System.ComponentModel.Composition;
using System.Windows.Media;

using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Utilities;

namespace AlekseyNagovitsyn.TfsPendingChangesMargin.Settings
{
    /// <summary>
    /// The editor format definition for <see cref="LineDiffType.Modified"/> line margin.
    /// </summary>
    [ClassificationType(ClassificationTypeNames = DefinitionKey)]
    [Name(DefinitionKey)]
    [DisplayName(DefinitionTitle)]
    [Order(Before = Priority.Default)]
    [TextViewRole(PredefinedTextViewRoles.Document)]
    [UserVisible(true)]
    [Export(typeof(EditorFormatDefinition))]
    [ContentType("text")]
    internal sealed class ModifiedLineMarginFormatDefinition : EditorFormatDefinition
    {
        /// <summary>
        /// The definition has been stored in <see cref="IEditorFormatMap"/> by this key.
        /// </summary>
        private const string DefinitionKey = EditorMargin.MarginName + "_ModifiedLineMargin";

        /// <summary>
        /// Caption of the definition in "Tools -> Environment -> Fonts and Colors" dialog.
        /// </summary>
        private const string DefinitionTitle = EditorMargin.MarginName + " Modified line margin";

        /// <summary>
        /// Initializes a new instance of the <see cref="ModifiedLineMarginFormatDefinition"/> class with default properties.
        /// </summary>
        public ModifiedLineMarginFormatDefinition()
        {
            DisplayName = DefinitionTitle;
            ForegroundColor = Colors.DodgerBlue;
            BackgroundCustomizable = false;
        }
    }
}