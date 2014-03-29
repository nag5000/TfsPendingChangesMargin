using System;
using System.Reflection;
using System.Windows;
using System.Windows.Media;

using Microsoft.VisualStudio.Text.Classification;

namespace AlekseyNagovitsyn.TfsPendingChangesMargin.Settings
{
    /// <summary>
    /// Extensions for <see cref="IEditorFormatMap"/>.
    /// </summary>
    internal static class EditorFormatMapExtensions
    {
        /// <summary>
        /// Get foreground brush from specified <see cref="EditorFormatDefinition"/> type.
        /// </summary>
        /// <typeparam name="T">Type of <see cref="EditorFormatDefinition"/>.</typeparam>
        /// <param name="formatMap">Map that contains format definitions.</param>
        /// <returns>The foreground brush.</returns>
        public static SolidColorBrush GetForeground<T>(this IEditorFormatMap formatMap)
            where T : EditorFormatDefinition
        {
            var attr = typeof(T).GetCustomAttribute<ClassificationTypeAttribute>();
            if (attr == null)
                throw new ArgumentException(string.Format("Type \"{0}\" does not contain \"{1}\" attribute.", typeof(T).FullName, typeof(ClassificationTypeAttribute).FullName));

            ResourceDictionary properties = formatMap.GetProperties(attr.ClassificationTypeNames);
            var brush = (SolidColorBrush)properties[EditorFormatDefinition.ForegroundBrushId];
            return brush;
        }
    }
}