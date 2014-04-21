using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Shapes;

using Microsoft.TeamFoundation.Diff;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Formatting;

namespace AlekseyNagovitsyn.TfsPendingChangesMargin
{
    /// <summary>
    /// TFS Pending Changes Margin for the Text Editor.
    /// </summary>
    internal class EditorMargin : Canvas, IWpfTextViewMargin
    {
        #region Fields

        /// <summary>
        /// The name of this margin.
        /// </summary>
        internal const string MarginName = "TfsPendingChangesMargin_EditorMargin";

        /// <summary>
        /// Left indent of the margin element.
        /// </summary>
        private const double MarginElementLeft = 0.5;

        /// <summary>
        /// Width of the margin element.
        /// </summary>
        private const double MarginElementWidth = 5;

        /// <summary>
        /// Right indent of the margin.
        /// </summary>
        /// <remarks>A little indent before the outline margin.</remarks>
        private const double MarginRightIndent = 2;

        /// <summary>
        /// The current instance of <see cref="IWpfTextView"/>.
        /// </summary>
        private readonly IWpfTextView _textView;

        /// <summary>
        /// The class which receives, processes and provides necessary data for <see cref="EditorMargin"/>.
        /// </summary>
        private readonly MarginCore _marginCore;

        /// <summary>
        /// The margin has been disposed of.
        /// </summary>
        private bool _isDisposed;

        #endregion Fields

        #region IWpfTextViewMargin Members

        /// <summary>
        /// The <see cref="System.Windows.FrameworkElement"/> that implements the visual representation of the margin.
        /// Since this margin implements <see cref="System.Windows.Controls.Canvas"/>, this is the object which renders the margin.
        /// </summary>
        public FrameworkElement VisualElement
        {
            get
            {
                ThrowIfDisposed();
                return this;
            }
        }

        /// <summary>
        /// Since this is a vertical margin, its height will be bound to the height of the text view. 
        /// </summary>
        public double MarginSize
        {
            get
            {
                ThrowIfDisposed();
                return ActualHeight;
            }
        }

        /// <summary>
        /// Determines whether the margin is enabled.
        /// </summary>
        public bool Enabled
        {
            get
            {
                ThrowIfDisposed();
                return true;
            }
        }

        /// <summary>
        /// Returns an instance of the margin if this is the margin that has been requested.
        /// </summary>
        /// <param name="marginName">The name of the margin requested.</param>
        /// <returns>An instance of <see cref="EditorMargin"/> or <c>null</c>.</returns>
        public ITextViewMargin GetTextViewMargin(string marginName)
        {
            return (marginName == MarginName) ? this : null;
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            if (!_isDisposed)
            {
                _marginCore.MarginRedraw -= OnMarginCoreMarginRedraw;
                _marginCore.ExceptionThrown -= OnMarginCoreExceptionThrown;
                _marginCore.Dispose();    

                GC.SuppressFinalize(this);
                _isDisposed = true;
            }
        }

        #endregion IWpfTextViewMargin Members

        /// <summary>
        /// Creates a <see cref="EditorMargin"/> for a given <see cref="IWpfTextView"/>.
        /// </summary>
        /// <param name="textView">The <see cref="IWpfTextView"/> to attach the margin to.</param>
        /// <param name="marginCore">The class which receives, processes and provides necessary data for <see cref="EditorMargin"/>.</param>
        public EditorMargin(IWpfTextView textView, MarginCore marginCore)
        {
            Debug.WriteLine("Entering constructor.", MarginName);

            if (!marginCore.IsEnabled)
                return;

            InitializeLayout();

            _textView = textView;
            _marginCore = marginCore;

            marginCore.MarginRedraw += OnMarginCoreMarginRedraw;
            marginCore.ExceptionThrown += OnMarginCoreExceptionThrown;

            if (marginCore.IsActivated)
                DrawMargins(marginCore.GetChangedLines());
        }

        /// <summary>
        /// Checks that the line intersects with one of the changed lines.
        /// </summary>
        /// <param name="line">The checked line.</param>
        /// <param name="changedLines">Collection of the changed lines.</param>
        /// <param name="intersectedLine">A line from the <see cref="changedLines"/> which was intersected with the <see cref="line"/>.</param>
        /// <returns>Returns <c>true</c>, if the line contains changes.</returns>
        /// <exception cref="ArgumentException">The supplied <see cref="ITextSnapshotLine"/> is on an incorrect snapshot.</exception>
        private static bool ContainsChanges(ITextViewLine line, IReadOnlyList<ITextSnapshotLine> changedLines, out ITextSnapshotLine intersectedLine)
        {
            int changedLineIndex = 0;
            int changedLinesCount = changedLines.Count;
            while (changedLineIndex < changedLinesCount)
            {
                int index = (changedLineIndex + changedLinesCount) / 2;
                if ((int)line.Start <= (int)changedLines[index].End)
                    changedLinesCount = index;
                else
                    changedLineIndex = index + 1;
            }

            if (changedLineIndex >= changedLines.Count)
            {
                intersectedLine = null;
                return false;
            }

            bool containsChanges;
            var checkedIntersectedLine = changedLines[changedLineIndex];

            if ((int)line.EndIncludingLineBreak != line.Snapshot.Length || line.LineBreakLength != 0)
                containsChanges = (int)line.EndIncludingLineBreak > (int)checkedIntersectedLine.Start;
            else
                containsChanges = (int)line.EndIncludingLineBreak >= (int)checkedIntersectedLine.Start;

            intersectedLine = containsChanges ? checkedIntersectedLine : null;
            return containsChanges;
        }

        /// <summary>
        /// Initialize layout.
        /// </summary>
        private void InitializeLayout()
        {
            Width = MarginElementWidth + MarginRightIndent;
            ClipToBounds = true;
        }

        /// <summary>
        /// Event handler that occurs when an unhandled exception was catched in <see cref="MarginCore"/>.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event arguments.</param>
        private void OnMarginCoreExceptionThrown(object sender, ExceptionThrownEventArgs e)
        {
            if (!e.Handled)
            {
                ShowException(e.Exception);
                e.Handled = true;
            }
        }

        /// <summary>
        /// Event handler that occurs when margins needs to be redrawn.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event arguments.</param>
        private void OnMarginCoreMarginRedraw(object sender, MarginRedrawEventArgs e)
        {
            switch (e.Reason)
            {
                case MarginDrawReason.InternalReason:
                case MarginDrawReason.VersionControlItemChanged:
                case MarginDrawReason.TextViewZoomLevelChanged:
                case MarginDrawReason.TextDocFileActionOccurred:
                case MarginDrawReason.TextViewTextChanged:
                case MarginDrawReason.TextViewLayoutChanged:
                case MarginDrawReason.EditorFormatMapChanged:
                    DrawMargins(e.DiffLines);
                    break;

                case MarginDrawReason.ScrollMapMappingChanged:
                    return;

                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        /// <summary>
        /// Draw margins for each diff line.
        /// </summary>
        /// <param name="diffLines">Differences between the current document and the version in TFS.</param>
        private void DrawMargins(DiffLinesCollection diffLines)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => DrawMargins(diffLines));
                return;
            }

            Children.Clear();

            foreach (ITextViewLine viewLine in _textView.TextViewLines)
            {
                DiffChangeType diffType;
                ITextSnapshotLine line;
                
                try
                {
                    if (ContainsChanges(viewLine, diffLines[DiffChangeType.Change], out line))
                    {
                        diffType = DiffChangeType.Change;
                    }
                    else if (ContainsChanges(viewLine, diffLines[DiffChangeType.Insert], out line))
                    {
                        diffType = DiffChangeType.Insert;

                        ITextSnapshotLine tmp;
                        if (ContainsChanges(viewLine, diffLines[DiffChangeType.Delete], out tmp))
                        {
                            diffType = DiffChangeType.Change;
                            line = tmp;
                        }
                    }
                    else if (ContainsChanges(viewLine, diffLines[DiffChangeType.Delete], out line))
                    {
                        diffType = DiffChangeType.Delete;
                    }
                    else
                    {
                        continue;
                    }
                }
                catch (ArgumentException)
                {
                    // The supplied line is on an incorrect snapshot (old version).
                    return;
                }

                IDiffChange diffChangeInfo = diffLines[line];

                var rect = new Rectangle { Height = viewLine.Height, Width = MarginElementWidth };
                SetLeft(rect, MarginElementLeft);
                SetTop(rect, viewLine.Top - _textView.ViewportTop);

                switch (diffType)
                {
                    case DiffChangeType.Insert:
                        rect.Fill = _marginCore.MarginSettings.AddedLineMarginBrush;
                        break;
                    case DiffChangeType.Change:
                        rect.Fill = _marginCore.MarginSettings.ModifiedLineMarginBrush;
                        break;
                    case DiffChangeType.Delete:
                        rect.Fill = _marginCore.MarginSettings.RemovedLineMarginBrush;
                        break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }

                if (diffType != DiffChangeType.Insert)
                {
                    rect.MouseEnter += (sender, args) =>
                    {
                        if (rect.ToolTip == null)
                        {
                            ToolTipService.SetShowDuration(rect, 3600000);
                            string text = _marginCore.GetOriginalText(diffChangeInfo.OriginalStart, diffChangeInfo.OriginalEnd);
                            rect.ToolTip = text;

                            // TODO: Если регион свернут, то тултип будет только для первого изменения в нем, т.к. ContainsChanges возвращает первую пересекаемую строку. 
                            // Все пересекаемые строки не учитываются. Нужно добавить поддержку всех пересекаемых строк, но только для показа тултипа, чтобы не замедлять рендеринг.
                        }
                    };
                }

                Children.Add(rect);
            }
        }

        /// <summary>
        /// Show an unhandled exception that was thrown by the margin.
        /// </summary>
        /// <param name="ex">The exception instance.</param>
        private void ShowException(Exception ex)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => ShowException(ex));
                return;
            }

            string msg = string.Format(
                "Unhandled exception was thrown in {0}.{2}" +
                "Please contact with developer. You can copy this message to the Clipboard with CTRL+C.{2}{2}" +
                "{1}",
                Properties.Resources.ProductName,
                ex,
                Environment.NewLine);

            MessageBox.Show(msg, Properties.Resources.ProductName, MessageBoxButton.OK, MessageBoxImage.Error);
        }

        /// <summary>
        /// Throw <see cref="ObjectDisposedException"/> if the margin was disposed.
        /// </summary>
        private void ThrowIfDisposed()
        {
            if (_isDisposed)
                throw new ObjectDisposedException(MarginName);
        }
    }
}
