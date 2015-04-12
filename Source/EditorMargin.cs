using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Shapes;

using Microsoft.TeamFoundation.Diff;
using Microsoft.TeamFoundation.VersionControl.Client;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Formatting;
using Microsoft.VisualStudio.Text.Operations;

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
        /// Contains undo transactions.
        /// </summary>
        private readonly ITextUndoHistory _undoHistory;

        /// <summary>
        /// The margin has been disposed of.
        /// </summary>
        private bool _isDisposed;

        /// <summary>
        /// Context menu for the margin element.
        /// </summary>
        private ContextMenu _contextMenu;

        /// <summary>
        /// Menu item for copying commited text to the Clipboard.
        /// </summary>
        private MenuItem _copyMenuItem;

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
        /// <param name="undoHistoryRegistry">Maintains the relationship between text buffers and <see cref="ITextUndoHistory"/> objects.</param>
        /// <param name="marginCore">The class which receives, processes and provides necessary data for <see cref="EditorMargin"/>.</param>
        public EditorMargin(IWpfTextView textView, ITextUndoHistoryRegistry undoHistoryRegistry, MarginCore marginCore)
        {
            Debug.WriteLine("Entering constructor.", MarginName);

            if (!marginCore.IsEnabled)
                return;

            InitializeLayout();

            _textView = textView;
            _marginCore = marginCore;

            undoHistoryRegistry.TryGetHistory(textView.TextBuffer, out _undoHistory);

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
                if (line.Start.Position <= changedLines[index].End.Position)
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

            if (line.EndIncludingLineBreak.Position != line.Snapshot.Length || line.LineBreakLength != 0)
                containsChanges = line.EndIncludingLineBreak.Position > checkedIntersectedLine.Start.Position;
            else
                containsChanges = line.EndIncludingLineBreak.Position >= checkedIntersectedLine.Start.Position;

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

            _contextMenu = new ContextMenu();

            _copyMenuItem = new MenuItem { Header = "Copy Commited Text" };
            _copyMenuItem.Click += CopyCommitedTextMenuItemOnClick;
            _contextMenu.Items.Add(_copyMenuItem);

            var rollbackChangeMenuItem = new MenuItem { Header = "Rollback" };
            rollbackChangeMenuItem.Click += RollbackChangeMenuItemOnClick;
            _contextMenu.Items.Add(rollbackChangeMenuItem);

            var rollbackAllButThisChangeMenuItem = new MenuItem { Header = "Rollback All But This" };
            rollbackAllButThisChangeMenuItem.Click += RollbackAllButThisChangeMenuItemOnClick;
            _contextMenu.Items.Add(rollbackAllButThisChangeMenuItem);

            _contextMenu.Items.Add(new Separator());

            var compareChangeRegionMenuItem = new MenuItem { Header = "Compare Region..." };
            compareChangeRegionMenuItem.Click += CompareChangeRegionMenuItemOnClick;
            _contextMenu.Items.Add(compareChangeRegionMenuItem);

            var compareDocumentMenuItem = new MenuItem { Header = "Compare Document..." };
            compareDocumentMenuItem.Click += CompareDocumentMenuItemOnClick;
            _contextMenu.Items.Add(compareDocumentMenuItem);
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
                case MarginDrawReason.GeneralSettingsChanged:
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

                var rect = new Rectangle
                {
                    Height = viewLine.Height, 
                    Width = MarginElementWidth,
                    Cursor = Cursors.Hand,
                    ContextMenu = _contextMenu,
                    Tag = new { DiffChangeInfo = diffChangeInfo, Line = line }
                };
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
                    rect.MouseEnter += OnMarginElementMouseEnter;

                rect.MouseLeftButtonDown += OnMarginElementMouseLeftButtonDown;
                rect.MouseLeftButtonUp += OnMarginElementMouseLeftButtonUp;

                Children.Add(rect);
            }
        }

        /// <summary>
        /// Event handler that occurs when "Copy commited text" menu item is clicked.
        /// </summary>
        /// <param name="sender">Event sender (the <see cref="MenuItem"/>).</param>
        /// <param name="routedEventArgs">Event arguments.</param>
        private void CopyCommitedTextMenuItemOnClick(object sender, RoutedEventArgs routedEventArgs)
        {
            try
            {
                var marginElement = (FrameworkElement)_contextMenu.PlacementTarget;
                dynamic data = marginElement.Tag;
                IDiffChange diffChange = data.DiffChangeInfo;

                CopyCommitedText(diffChange);

                routedEventArgs.Handled = true;
            }
            catch (Exception ex)
            {
                ShowException(ex);
            }
        }

        /// <summary>
        /// Event handler that occurs when "Rollback modified text" menu item is clicked.
        /// </summary>
        /// <param name="sender">Event sender (the <see cref="MenuItem"/>).</param>
        /// <param name="routedEventArgs">Event arguments.</param>
        private void RollbackChangeMenuItemOnClick(object sender, RoutedEventArgs routedEventArgs)
        {
            try
            {
                var marginElement = (FrameworkElement)_contextMenu.PlacementTarget;
                dynamic data = marginElement.Tag;
                IDiffChange diffChange = data.DiffChangeInfo;

                RollbackChange(diffChange);

                routedEventArgs.Handled = true;
            }
            catch (Exception ex)
            {
                ShowException(ex);
            }
        }

        /// <summary>
        /// Event handler that occurs when "Rollback All But This" menu item is clicked.
        /// </summary>
        /// <param name="sender">Event sender (the <see cref="MenuItem"/>).</param>
        /// <param name="routedEventArgs">Event arguments.</param>
        private void RollbackAllButThisChangeMenuItemOnClick(object sender, RoutedEventArgs routedEventArgs)
        {
            try
            {
                var marginElement = (FrameworkElement)_contextMenu.PlacementTarget;
                dynamic data = marginElement.Tag;
                IDiffChange diffChange = data.DiffChangeInfo;

                RollbackAllButThisChange(diffChange);

                routedEventArgs.Handled = true;
            }
            catch (Exception ex)
            {
                ShowException(ex);
            }
        }

        /// <summary>
        /// Event handler that occurs when "Compare region with diff tool" menu item is clicked.
        /// </summary>
        /// <param name="sender">Event sender (the <see cref="MenuItem"/>).</param>
        /// <param name="routedEventArgs">Event arguments.</param>
        private void CompareChangeRegionMenuItemOnClick(object sender, RoutedEventArgs routedEventArgs)
        {
            try
            {
                var marginElement = (FrameworkElement)_contextMenu.PlacementTarget;
                dynamic data = marginElement.Tag;
                IDiffChange diffChange = data.DiffChangeInfo;

                CompareChangeRegion(diffChange);

                routedEventArgs.Handled = true;
            }
            catch (Exception ex)
            {
                ShowException(ex);
            }
        }

        /// <summary>
        /// Event handler that occurs when "Compare document with latest version" menu item is clicked.
        /// </summary>
        /// <param name="sender">Event sender (the <see cref="MenuItem"/>).</param>
        /// <param name="routedEventArgs">Event arguments.</param>
        private void CompareDocumentMenuItemOnClick(object sender, RoutedEventArgs routedEventArgs)
        {
            try
            {
                CompareDocumentWithLatestVersion();
                routedEventArgs.Handled = true;
            }
            catch (Exception ex)
            {
                ShowException(ex);
            }
        }

        /// <summary>
        /// Event handler that occurs when the mouse pointer enters the bounds of the margin element. 
        /// </summary>
        /// <param name="sender">Event sender (the margin element).</param>
        /// <param name="args">Event arguments.</param>
        private void OnMarginElementMouseEnter(object sender, MouseEventArgs args)
        {
            var marginElement = (FrameworkElement)sender;
            dynamic data = marginElement.Tag;
            IDiffChange diffChange = data.DiffChangeInfo;

            if (marginElement.ToolTip == null)
            {
                ToolTipService.SetShowDuration(marginElement, 3600000);
                string text = _marginCore.GetOriginalText(diffChange, true);
                marginElement.ToolTip = text;

                // TODO: Если регион свернут, то тултип будет только для первого изменения в нем, т.к. ContainsChanges возвращает первую пересекаемую строку. 
                // Все пересекаемые строки не учитываются. Нужно добавить поддержку всех пересекаемых строк, но только для показа тултипа, чтобы не замедлять рендеринг.
            }
        }

        /// <summary>
        /// Event handler that occurs when the left mouse button is pressed while the mouse pointer is over the margin element. 
        /// </summary>
        /// <param name="sender">Event sender (the margin element).</param>
        /// <param name="args">Event arguments.</param>
        private void OnMarginElementMouseLeftButtonDown(object sender, MouseButtonEventArgs args)
        {
            args.Handled = true;
        }

        /// <summary>
        /// Event handler that occurs when the left mouse button is released while the mouse pointer is over the margin element. 
        /// </summary>
        /// <param name="sender">Event sender (the margin element).</param>
        /// <param name="args">Event arguments.</param>
        private void OnMarginElementMouseLeftButtonUp(object sender, MouseButtonEventArgs args)
        {
            var element = (FrameworkElement)sender;

            // Disable copying commited text if diff inserted (otherwise an empty string is always copied).
            dynamic data = element.Tag;
            IDiffChange diffChange = data.DiffChangeInfo;
            _copyMenuItem.IsEnabled = diffChange.ChangeType != DiffChangeType.Insert;

            _contextMenu.PlacementTarget = element;
            _contextMenu.IsOpen = true;
            args.Handled = true;
        }

        /// <summary>
        /// Copy commited text to the Clipboard.
        /// </summary>
        /// <param name="diffChange">Information about a specific difference between two sequences.</param>
        private void CopyCommitedText(IDiffChange diffChange)
        {
            string text = _marginCore.GetOriginalText(diffChange, true);
            Clipboard.SetText(text, TextDataFormat.UnicodeText);
        }

        /// <summary>
        /// Rollback modified text.
        /// </summary>
        /// <param name="diffChange">Information about a specific difference between two sequences.</param>
        private void RollbackChange(IDiffChange diffChange)
        {
            ITextEdit edit = _textView.TextBuffer.CreateEdit();

            try
            {
                ITextSnapshot snapshot = edit.Snapshot;
                ITextSnapshotLine startLine = snapshot.GetLineFromLineNumber(diffChange.ModifiedStart);
                ITextSnapshotLine endLine = snapshot.GetLineFromLineNumber(diffChange.ModifiedEnd);

                int start;
                if (diffChange.ChangeType != DiffChangeType.Delete)
                {
                    start = startLine.Start.Position;
                    int length = endLine.EndIncludingLineBreak.Position - start;
                    edit.Delete(start, length);
                }
                else
                {
                    if (startLine.LineNumber == 0 && endLine.LineNumber == 0)
                        start = startLine.Start.Position;
                    else
                        start = startLine.EndIncludingLineBreak.Position;
                }

                if (diffChange.ChangeType != DiffChangeType.Insert)
                {
                    string text = _marginCore.GetOriginalText(diffChange, false);
                    edit.Insert(start, text);
                }

                ApplyEdit(edit, "Rollback Modified Region");
            }
            catch (Exception)
            {
                edit.Cancel();
                throw;
            }
        }

        /// <summary>
        /// Rollback all modified text in the document except specified <see cref="IDiffChange"/>.
        /// </summary>
        /// <param name="diffChange">Information about a specific difference between two sequences.</param>
        private void RollbackAllButThisChange(IDiffChange diffChange)
        {
            ITextEdit edit = _textView.TextBuffer.CreateEdit();
            Span viewSpan;

            try
            {
                string modifiedRegionText = _marginCore.GetModifiedText(diffChange, false);
                string originalText = _marginCore.GetOriginalText();

                edit.Delete(0, edit.Snapshot.Length);
                edit.Insert(0, originalText);
                ApplyEdit(edit, "Undo Modified Text");

                edit = _textView.TextBuffer.CreateEdit();

                ITextSnapshotLine startLine = edit.Snapshot.GetLineFromLineNumber(diffChange.OriginalStart);
                ITextSnapshotLine endLine = edit.Snapshot.GetLineFromLineNumber(diffChange.OriginalEnd);
                int start = startLine.Start.Position;
                int length = endLine.EndIncludingLineBreak.Position - start;

                switch (diffChange.ChangeType)
                {
                    case DiffChangeType.Insert:
                        edit.Insert(start, modifiedRegionText);
                        viewSpan = new Span(start, modifiedRegionText.Length);
                        break;

                    case DiffChangeType.Delete:
                        edit.Delete(start, length);
                        viewSpan = new Span(start, 0);
                        break;

                    case DiffChangeType.Change:
                        edit.Replace(start, length, modifiedRegionText);
                        viewSpan = new Span(start, modifiedRegionText.Length);
                        break;

                    default:
                        throw new InvalidEnumArgumentException();
                }

                ApplyEdit(edit, "Restore Modified Region");
            }
            catch (Exception)
            {
                edit.Cancel();
                throw;
            }

            var viewSnapshotSpan = new SnapshotSpan(_textView.TextSnapshot, viewSpan);
            _textView.ViewScroller.EnsureSpanVisible(viewSnapshotSpan, EnsureSpanVisibleOptions.AlwaysCenter);
        }

        /// <summary>
        /// Commits all modifications made with specified <see cref="ITextBufferEdit"/> and link the commit 
        /// with <see cref="ITextUndoTransaction"/> in the Editor, if <see cref="_undoHistory"/> is available.
        /// </summary>
        /// <param name="edit">A set of editing operations on an <see cref="ITextBuffer"/>.</param>
        /// <param name="description">The description of the transaction.</param>
        private void ApplyEdit(ITextBufferEdit edit, string description)
        {
            if (_undoHistory != null)
            {
                using (ITextUndoTransaction transaction = _undoHistory.CreateTransaction(description))
                {
                    edit.Apply();
                    transaction.Complete();
                }
            }
            else
            {
                edit.Apply();
            }
        }

        /// <summary>
        /// Visual compare current text document with the latest version in Version Control.
        /// </summary>
        private void CompareDocumentWithLatestVersion()
        {
            Item item = _marginCore.VersionControlItem;
            VersionControlServer vcs = item.VersionControlServer;
            ITextDocument textDoc = _marginCore.TextDocument;

            IDiffItem source = Difference.CreateTargetDiffItem(vcs, item.ServerItem, VersionSpec.Latest, 0, VersionSpec.Latest);
            var target = new DiffItemLocalFile(textDoc.FilePath, textDoc.Encoding.CodePage, textDoc.LastContentModifiedTime, false);
            Difference.VisualDiffItems(vcs, source, target);
        }

        /// <summary>
        /// Visual compare <see cref="IDiffChange"/> with the Visual Studio Diff Tool.
        /// </summary>
        /// <param name="diffChange">Information about a specific difference between two sequences.</param>
        private void CompareChangeRegion(IDiffChange diffChange)
        {
            const string SourceFileTag = "Server";
            const string TargetFileTag = "Local";

            const bool IsSourceReadOnly = true;
            const bool IsTargetReadOnly = true;

            const bool DeleteSourceOnExit = true;
            const bool DeleteTargetOnExit = true;

            const string FileLabelTemplateSingleLine = "{0};{1}[line {2}]";
            const string FileLabelTemplateBetweenLines = "{0};{1}[between lines {3}-{2}]";
            const string FileLabelTemplateLinesRange = "{0};{1}[lines {2}-{3}]";

            string sourceFileLabelTemplate;
            if (diffChange.OriginalStart == diffChange.OriginalEnd)
                sourceFileLabelTemplate = FileLabelTemplateSingleLine;
            else if (diffChange.ChangeType == DiffChangeType.Insert)
                sourceFileLabelTemplate = FileLabelTemplateBetweenLines;
            else
                sourceFileLabelTemplate = FileLabelTemplateLinesRange;

            string sourceFileLabel = string.Format(
                    sourceFileLabelTemplate,
                    _marginCore.VersionControlItem.ServerItem,
                    "T;" /* Latest version token */,
                    diffChange.OriginalStart + 1,
                    diffChange.OriginalEnd + 1);

            string targetFileLabelTemplate;
            if (diffChange.ModifiedStart == diffChange.ModifiedEnd)
                targetFileLabelTemplate = FileLabelTemplateSingleLine;
            else if (diffChange.ChangeType == DiffChangeType.Delete)
                targetFileLabelTemplate = FileLabelTemplateBetweenLines;
            else
                targetFileLabelTemplate = FileLabelTemplateLinesRange;

            string targetFileLabel = string.Format(
                    targetFileLabelTemplate,
                    _marginCore.TextDocument.FilePath,
                    string.Empty /* No additional parameter */,
                    diffChange.ModifiedStart + 1,
                    diffChange.ModifiedEnd + 1);

            const bool RemoveLastLineTerminator = true;

            string sourceText = _marginCore.GetOriginalText(diffChange, RemoveLastLineTerminator);
            string sourceFilePath = System.IO.Path.GetTempFileName();
            if (!string.IsNullOrEmpty(sourceText))
                File.WriteAllText(sourceFilePath, sourceText);

            string targetText = _marginCore.GetModifiedText(diffChange, RemoveLastLineTerminator);
            string targetFilePath = System.IO.Path.GetTempFileName();
            if (!string.IsNullOrEmpty(targetText))
                File.WriteAllText(targetFilePath, targetText);

            Difference.VisualDiffFiles(
                sourceFilePath,
                targetFilePath,
                SourceFileTag,
                TargetFileTag,
                sourceFileLabel,
                targetFileLabel,
                IsSourceReadOnly,
                IsTargetReadOnly,
                DeleteSourceOnExit,
                DeleteTargetOnExit);
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
