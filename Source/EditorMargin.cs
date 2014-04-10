using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Shapes;

using Microsoft.TeamFoundation.Diff;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Formatting;
using Microsoft.VisualStudio.Text.Outlining;

namespace AlekseyNagovitsyn.TfsPendingChangesMargin
{
    /// <summary>
    /// A class detailing the margin's visual definition including both size and content.
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
        /// Provides outlining functionality.
        /// </summary>
        private readonly IOutliningManager _outliningManager;

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
        /// <param name="outliningManagerService">Service that provides the <see cref="IOutliningManager"/>.</param>
        /// <param name="marginCore">The class which receives, processes and provides necessary data for <see cref="EditorMargin"/>.</param>
        public EditorMargin(IWpfTextView textView, IOutliningManagerService outliningManagerService, MarginCore marginCore)
        {
            Debug.WriteLine("Entering constructor.", MarginName);

            if (!marginCore.IsEnabled)
                return;

            InitializeLayout();

            _textView = textView;
            _marginCore = marginCore;
            _outliningManager = outliningManagerService.GetOutliningManager(textView);

            marginCore.MarginRedraw += OnMarginCoreMarginRedraw;
            marginCore.ExceptionThrown += OnMarginCoreExceptionThrown;

            if (marginCore.IsActivated)
                marginCore.RaiseMarginRedraw();
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
            DrawMargins(e.DiffLines);
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
        /// Draw margins for each diff line.
        /// </summary>
        /// <param name="diffLines">Differences between the current document and the version in TFS.</param>
        private void DrawMargins(Dictionary<ITextSnapshotLine, DiffChangeType> diffLines)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => DrawMargins(diffLines));
                return;
            }

            Debug.Assert(_textView != null, "_textView is null.");
            Debug.Assert(_outliningManager != null, "_outliningManager is null.");

            Children.Clear();

            var rectMap = new Dictionary<double, KeyValuePair<DiffChangeType, Rectangle>>();
            foreach (KeyValuePair<ITextSnapshotLine, DiffChangeType> diffLine in diffLines)
            {
                ITextSnapshotLine line = diffLine.Key;
                DiffChangeType diffType = diffLine.Value;

                IWpfTextViewLine viewLine;

                try
                {
                    viewLine = _textView.GetTextViewLineContainingBufferPosition(line.Start);
                    Debug.Assert(viewLine != null, "viewLine is null.");
                }
                catch (InvalidOperationException)
                {
                    if (_textView.IsClosed)
                        return;

                    throw;
                }
                catch (ArgumentException)
                {
                    // The supplied SnapshotPoint is on an incorrect snapshot (old version).
                    return;
                }

                switch (viewLine.VisibilityState)
                {
                    case VisibilityState.Unattached:
                    case VisibilityState.Hidden:
                        continue;

                    case VisibilityState.PartiallyVisible:
                    case VisibilityState.FullyVisible:
                        break;

                    default:
                        throw new ArgumentOutOfRangeException();
                }

                if (rectMap.ContainsKey(viewLine.Top))
                {
                    #if DEBUG
                    IEnumerable<ICollapsed> collapsedRegions = _outliningManager.GetCollapsedRegions(viewLine.Extent, true);
                    bool lineIsCollapsed = collapsedRegions.Any();
                    Debug.Assert(lineIsCollapsed, "line should be collapsed.");
                    #endif

                    KeyValuePair<DiffChangeType, Rectangle> rectMapValue = rectMap[viewLine.Top];
                    if (rectMapValue.Key != DiffChangeType.Change && rectMapValue.Key != diffType)
                    {
                        rectMapValue.Value.Fill = _marginCore.MarginSettings.ModifiedLineMarginBrush;
                        rectMap[viewLine.Top] = new KeyValuePair<DiffChangeType, Rectangle>(DiffChangeType.Change, rectMapValue.Value);
                    }

                    continue;
                }

                var rect = new Rectangle { Height = viewLine.Height, Width = MarginElementWidth };
                SetLeft(rect, MarginElementLeft);
                SetTop(rect, viewLine.Top - _textView.ViewportTop);
                rectMap.Add(viewLine.Top, new KeyValuePair<DiffChangeType, Rectangle>(diffType, rect));

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
