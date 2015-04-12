using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Shapes;

using Microsoft.TeamFoundation.Diff;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;

namespace AlekseyNagovitsyn.TfsPendingChangesMargin
{
    /// <summary>
    /// TFS Pending Changes Margin on the Enhanced Scrollbar.
    /// </summary>
    internal class ScrollbarMargin : Canvas, IWpfTextViewMargin
    {
        #region Fields

        /// <summary>
        /// The name of this margin.
        /// </summary>
        internal const string MarginName = "TfsPendingChangesMargin_ScrollbarMargin";

        /// <summary>
        /// A horizontal offset of the margin element.
        /// </summary>
        private const double MarginElementOffset = 1.0;

        /// <summary>
        /// The current instance of <see cref="IWpfTextView"/>.
        /// </summary>
        private readonly IWpfTextView _textView;

        /// <summary>
        /// The class which receives, processes and provides necessary data for <see cref="EditorMargin"/>.
        /// </summary>
        private readonly MarginCore _marginCore;

        /// <summary>
        /// A vertical scroll bar in the Text Editor.
        /// </summary>
        private readonly IVerticalScrollBar _scrollBar;

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
                _marginCore.Dispose();  

                GC.SuppressFinalize(this);
                _isDisposed = true;
            }
        }

        #endregion IWpfTextViewMargin Members

        /// <summary>
        /// Creates a <see cref="ScrollbarMargin"/> for a given <see cref="IWpfTextView"/>.
        /// </summary>
        /// <param name="textView">The <see cref="IWpfTextView"/> to attach the margin to.</param>
        /// <param name="marginContainer">Margin container. Is defined in the <see cref="ScrollbarMarginFactory"/> by the <see cref="MarginContainerAttribute"/>.</param>
        /// <param name="marginCore">The class which receives, processes and provides necessary data for <see cref="ScrollbarMargin"/>.</param>
        public ScrollbarMargin(IWpfTextView textView, IWpfTextViewMargin marginContainer, MarginCore marginCore)
        {
            Debug.WriteLine("Entering constructor.", MarginName);

            _textView = textView;
            _marginCore = marginCore;

            InitializeLayout();

            ITextViewMargin scrollBarMargin = marginContainer.GetTextViewMargin(PredefinedMarginNames.VerticalScrollBar);
            // ReSharper disable once SuspiciousTypeConversion.Global - scrollBarMargin is IVerticalScrollBar.
            _scrollBar = (IVerticalScrollBar)scrollBarMargin;

            marginCore.MarginRedraw += OnMarginCoreMarginRedraw;

            if (marginCore.IsActivated)
                DrawMargins(marginCore.GetChangedLines());
        }

        /// <summary>
        /// Initialize layout.
        /// </summary>
        private void InitializeLayout()
        {
            Width = _textView.Options.GetOptionValue(DefaultTextViewHostOptions.ChangeTrackingMarginWidthOptionId);
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
                case MarginDrawReason.EditorFormatMapChanged:
                case MarginDrawReason.GeneralSettingsChanged:
                case MarginDrawReason.ScrollMapMappingChanged:
                    DrawMargins(e.DiffLines);
                    break;

                case MarginDrawReason.TextViewLayoutChanged:
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
            
            foreach (DiffLineEntry diffLine in diffLines)
            {
                double top, bottom;
                try
                {
                    MapLineToPixels(diffLine.Line, out top, out bottom);
                }
                catch (ArgumentException)
                {
                    // The supplied line is on an incorrect snapshot (old version).
                    return;
                }

                var rect = new Rectangle
                {
                    Height = bottom - top,
                    Width = Width - MarginElementOffset,
                    Focusable = false,
                    IsHitTestVisible = false
                };
                SetLeft(rect, MarginElementOffset);
                SetTop(rect, top);

                switch (diffLine.ChangeType)
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
        /// Get top and bottom of a line for the scrollbar.
        /// </summary>
        /// <param name="line">A line of text from an <see cref="ITextSnapshot"/>.</param>
        /// <param name="top">Top coordinate for the scrollbar.</param>
        /// <param name="bottom">Bottom coordinate for the scrollbar.</param>
        /// <exception cref="ArgumentException">The supplied <see cref="ITextSnapshotLine"/> is on an incorrect snapshot.</exception>
        private void MapLineToPixels(ITextSnapshotLine line, out double top, out double bottom)
        {
            double mapTop = _scrollBar.Map.GetCoordinateAtBufferPosition(line.Start) - 0.5;
            double mapBottom = _scrollBar.Map.GetCoordinateAtBufferPosition(line.End) + 0.5;
            top = Math.Round(_scrollBar.GetYCoordinateOfScrollMapPosition(mapTop)) - 2.0;
            bottom = Math.Round(_scrollBar.GetYCoordinateOfScrollMapPosition(mapBottom)) + 2.0;
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