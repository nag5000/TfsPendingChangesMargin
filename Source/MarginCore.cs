using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using AlekseyNagovitsyn.TfsPendingChangesMargin.Settings;

using EnvDTE;
using EnvDTE80;

using Microsoft.TeamFoundation;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Diff;
using Microsoft.TeamFoundation.VersionControl.Client;
using Microsoft.TeamFoundation.VersionControl.Common;

using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.TeamFoundation;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Editor;

namespace AlekseyNagovitsyn.TfsPendingChangesMargin
{
    /// <summary>
    /// The class which receives, processes and provides necessary data for <see cref="EditorMargin"/>.
    /// </summary>
    internal sealed class MarginCore : IDisposable
    {
        #region Fields

        /// <summary>
        /// The current instance of <see cref="IWpfTextView"/>.
        /// </summary>
        private readonly IWpfTextView _textView;

        /// <summary>
        /// The current instance of <see cref="ITextDocument"/> which is associated with <see cref="_textView"/>.
        /// </summary>
        private readonly ITextDocument _textDoc;

        /// <summary>
        /// Map that contains properties from "Tools -> Environment -> Fonts and Colors" dialog in Visual Studio.
        /// </summary>
        private readonly IEditorFormatMap _formatMap;

        /// <summary>
        /// Represents the Team Foundation Server extensibility model within Visual Studio.
        /// </summary>
        private readonly TeamFoundationServerExt _tfExt;

        /// <summary>
        /// The margin settings.
        /// </summary>
        private readonly MarginSettings _marginSettings;

        /// <summary>
        /// Defines the mapping between character positions and scrollmap coordinates.
        /// </summary>
        private readonly IScrollMap _scrollMap;

        /// <summary>
        /// Lock-object used to synchronize <see cref="GetDifference"/> method.
        /// </summary>
        private readonly object _differenceLockObject = new object();

        /// <summary>
        /// Lock-object used to synchronize the margin drawing.
        /// </summary>
        private readonly object _drawLockObject = new object();

        /// <summary>
        /// The margin is enabled.
        /// </summary>
        private readonly bool _isEnabled = true;

        /// <summary>
        /// The margin has been disposed of.
        /// </summary>
        private bool _isDisposed;

        /// <summary>
        /// The margin is activated.
        /// </summary>
        private bool _isActivated;

        /// <summary>
        /// Represents the version control repository.
        /// </summary>
        private VersionControlServer _versionControl;

        /// <summary>
        /// Represents a committed version of the <see cref="_textDoc"/> in the version control server.
        /// </summary>
        private Item _versionControlItem;

        /// <summary>
        /// Content of the commited version of the <see cref="_textDoc"/> in the version control server.
        /// </summary>
        private Stream _versionControlItemStream;

        /// <summary>
        /// Task for observation over VersionControlItem up-dating.
        /// </summary>
        private Task _versionControlItemWatcher;

        /// <summary>
        /// A cancellation token source for task <see cref="_versionControlItemWatcher"/>.
        /// </summary>
        private CancellationTokenSource _versionControlItemWatcherCts;

        /// <summary>
        /// Collection that contains the result of comparing the document's local file with his source control version.
        /// </summary>
        private DiffLinesCollection _cachedChangedLines = new DiffLinesCollection();

        #endregion Fields

        /// <summary>
        /// Gets the margin settings.
        /// </summary>
        public MarginSettings MarginSettings
        {
            get { return _marginSettings; }
        }

        /// <summary>
        /// Determines whether the margin is enabled.
        /// </summary>
        public bool IsEnabled
        {
            get { return _isEnabled; }
        }

        /// <summary>
        /// Gets the margin is activated.
        /// </summary>
        public bool IsActivated
        {
            get { return _isActivated; }
        }

        /// <summary>
        /// Gets cached differences between the current document and the version in TFS.
        /// </summary>
        /// <returns>
        /// Collection that contains the result of comparing the document's local file with his source control version.
        /// </returns>
        public DiffLinesCollection GetChangedLines()
        {
            return _cachedChangedLines;
        }

        /// <summary>
        /// Event that raised when margins needs to be redrawn: 
        /// differences between text document and TFS item were changed or document layout was changed.
        /// </summary>
        public event EventHandler<MarginRedrawEventArgs> MarginRedraw;

        /// <summary>
        /// Event that raised when an unhandled exception was catched.
        /// </summary>
        public event EventHandler<ExceptionThrownEventArgs> ExceptionThrown;

        /// <summary>
        /// Creates a <see cref="MarginCore"/> for a given <see cref="IWpfTextView"/>.
        /// </summary>
        /// <param name="textView">The <see cref="IWpfTextView"/> to attach the margin to.</param>
        /// <param name="textDocumentFactoryService">Service that creates, loads, and disposes text documents.</param>
        /// <param name="vsServiceProvider">Visual Studio service provider.</param>
        /// <param name="formatMapService">Service that provides the <see cref="IEditorFormatMap"/>.</param>
        /// <param name="scrollMapFactoryService">Factory that creates or reuses an <see cref="IScrollMap"/> for an <see cref="ITextView"/>.</param>
        public MarginCore(IWpfTextView textView, ITextDocumentFactoryService textDocumentFactoryService, SVsServiceProvider vsServiceProvider, IEditorFormatMapService formatMapService, IScrollMapFactoryService scrollMapFactoryService)
        {
            Debug.WriteLine("Entering constructor.", Properties.Resources.ProductName);

            _textView = textView;
            if (!textDocumentFactoryService.TryGetTextDocument(_textView.TextBuffer, out _textDoc))
            {
                Debug.WriteLine("Can not retrieve TextDocument. Margin is disabled.", Properties.Resources.ProductName);
                _isEnabled = false;
                return;
            }

            _formatMap = formatMapService.GetEditorFormatMap(textView);
            _marginSettings = new MarginSettings(_formatMap);

            _scrollMap = scrollMapFactoryService.Create(textView);

            var dte = (DTE2)vsServiceProvider.GetService(typeof(DTE));
            _tfExt = dte.GetObject(typeof(TeamFoundationServerExt).FullName);
            Debug.Assert(_tfExt != null, "_tfExt is null.");
            _tfExt.ProjectContextChanged += OnTfExtProjectContextChanged;

            UpdateMargin();
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            if (_isDisposed)
                return;

            _isDisposed = true;

            SetMarginActivated(false);

            _tfExt.ProjectContextChanged -= OnTfExtProjectContextChanged;
            _cachedChangedLines.Clear();
            _versionControlItemStream.Dispose();
        }

        /// <summary>
        /// <see cref="DiffUtil.Diff"/> behaves incorrectly if the stream terminates in blank line - in that case it isn't considered. 
        /// And as a result, changes are calculated incorrectly. 
        /// This method is called for both compared streams before comparing and adds a nonblank line to them.
        /// </summary>
        /// <param name="stream">Stream to which bytes will be added.</param>
        /// <param name="encoding">Stream encoding.</param>
        private static void AppendShiftTokenToStream(Stream stream, Encoding encoding)
        {
            stream.Seek(0, SeekOrigin.End);
            byte[] eofShiftBytes = encoding.GetBytes(Environment.NewLine + "0");
            stream.Write(eofShiftBytes, 0, eofShiftBytes.Length);
        }

        /// <summary>
        /// Checks whether there are any changes between two <see cref="ITextBuffer"/> versions.
        /// </summary>
        /// <param name="oldVersion">An old text version.</param>
        /// <param name="currentVersion">The current text version.</param>
        /// <returns>Returns <c>true</c> if there are changes between two versions.</returns>
        private static bool AnyTextChanges(ITextVersion oldVersion, ITextVersion currentVersion)
        {
            for (; oldVersion != currentVersion; oldVersion = oldVersion.Next)
            {
                if (oldVersion.Changes.Count > 0)
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Adds a line of text from an <see cref="ITextSnapshot"/> to <see cref="DiffLinesCollection"/>.
        /// </summary>
        /// <param name="collection">Collection of differences between the current document and the version in TFS.</param>
        /// <param name="textSnapshot">The text snapshot.</param>
        /// <param name="lineNumber">The line number.</param>
        /// <param name="lineDiffType">Difference type of the line.</param>
        private static void AddLineToDiffLinesCollection(DiffLinesCollection collection, ITextSnapshot textSnapshot, int lineNumber, DiffChangeType lineDiffType)
        {
            if (lineNumber == -1 && lineDiffType == DiffChangeType.Delete)
                lineNumber = 0;

            ITextSnapshotLine line;
            try
            {
                line = textSnapshot.GetLineFromLineNumber(lineNumber);
            }
            catch (ArgumentOutOfRangeException ex)
            {
                string msg = string.Format("Line number {0} is out of range [0..{1}].", lineNumber, textSnapshot.LineCount);
                throw new ArgumentOutOfRangeException(msg, ex);
            }

            collection[lineDiffType].Add(line);
        }

        /// <summary>
        /// Raises the <see cref="MarginRedraw"/> event.
        /// </summary>
        /// <param name="reason">The reason of redrawing.</param>
        private void RaiseMarginRedraw(MarginDrawReason reason)
        {
            var eventHandler = MarginRedraw;
            if (eventHandler != null)
                eventHandler(this, new MarginRedrawEventArgs(_cachedChangedLines, reason));
        }

        /// <summary>
        /// Raises the <see cref="ExceptionThrown"/> event.
        /// </summary>
        /// <param name="exception">The exception that was thrown.</param>
        private void RaiseExceptionThrown(Exception exception)
        {
            var eventHandler = ExceptionThrown;
            var eventArgs = new ExceptionThrownEventArgs(exception);
            if (eventHandler != null)
                eventHandler(this, eventArgs);

            if (!eventArgs.Handled)
                throw exception;
        }

        /// <summary>
        /// Sets the activated state of the margin.
        /// </summary>
        /// <param name="marginIsActivated">The new activated state to assign to the margin.</param>
        private void SetMarginActivated(bool marginIsActivated)
        {
            if (_isActivated == marginIsActivated)
                return;

            _isActivated = marginIsActivated;
            Debug.WriteLine(string.Format("MarginActivity: {0} ({1})", marginIsActivated, _textDoc.FilePath), Properties.Resources.ProductName);

            if (marginIsActivated)
            {
                _textView.LayoutChanged += OnTextViewLayoutChanged;
                _textView.ZoomLevelChanged += OnTextViewZoomLevelChanged;
                _textDoc.FileActionOccurred += OnTextDocFileActionOccurred;
                _formatMap.FormatMappingChanged += OnFormatMapFormatMappingChanged;
                _versionControl.CommitCheckin += OnVersionControlCommitCheckin;
                _scrollMap.MappingChanged += OnScrollMapMappingChanged;

                _versionControlItemWatcherCts = new CancellationTokenSource();
                CancellationToken token = _versionControlItemWatcherCts.Token;
                _versionControlItemWatcher = new Task(ObserveVersionControlItem, token, token);
                _versionControlItemWatcher.Start();
            }
            else
            {
                _textView.LayoutChanged -= OnTextViewLayoutChanged;
                _textView.ZoomLevelChanged -= OnTextViewZoomLevelChanged;
                _textDoc.FileActionOccurred -= OnTextDocFileActionOccurred;
                _formatMap.FormatMappingChanged -= OnFormatMapFormatMappingChanged;
                _versionControl.CommitCheckin -= OnVersionControlCommitCheckin;
                _scrollMap.MappingChanged -= OnScrollMapMappingChanged;

                if (_versionControlItemWatcherCts != null)
                    _versionControlItemWatcherCts.Cancel();
            }
        }

        /// <summary>
        /// Refresh margin's data and redraw it.
        /// </summary>
        private void UpdateMargin()
        {
            var task = new Task(() =>
            {
                lock (_drawLockObject)
                {
                    bool success = RefreshVersionControl();
                    SetMarginActivated(success);
                    Redraw(false, MarginDrawReason.InternalReason);
                }
            });

            task.Start();
        }

        /// <summary>
        /// Refresh version control status and update associated fields.
        /// </summary>
        /// <returns>
        /// Returns <c>false</c>, if no connection to a team project.
        /// Returns <c>false</c>, if current text document (file) not associated with version control.
        /// Otherwise, returns <c>true</c>.
        /// </returns>
        private bool RefreshVersionControl()
        {
            string tfsServerUriString = _tfExt.ActiveProjectContext.DomainUri;
            if (string.IsNullOrEmpty(tfsServerUriString))
            {
                _versionControl = null;
                _versionControlItem = null;
                return false;
            }

            TfsTeamProjectCollection tfsProjCollections = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri(tfsServerUriString));
            _versionControl = (VersionControlServer)tfsProjCollections.GetService(typeof(VersionControlServer));
            try
            {
                _versionControlItem = GetVersionControlItem();
            }
            catch (VersionControlItemNotFoundException)
            {
                _versionControlItem = null;
                return false;
            }
            catch (TeamFoundationServiceUnavailableException)
            {
                _versionControlItem = null;
                return false;
            }

            DownloadVersionControlItem();
            return true;
        }

        /// <summary>
        /// Observation over VersionControlItem up-dating at regular intervals.
        /// </summary>
        /// <param name="cancellationTokenObject">A cancellation token for this method.</param>
        private void ObserveVersionControlItem(object cancellationTokenObject)
        {
            var cancellationToken = (CancellationToken)cancellationTokenObject;

            while (true)
            {
                try
                {
                    const int VersionControlItemObservationInterval = 30000;
                    System.Threading.Thread.Sleep(VersionControlItemObservationInterval);

                    if (cancellationToken.IsCancellationRequested)
                        break;

                    lock (_drawLockObject)
                    {
                        Item versionControlItem;
                        try
                        {
                            versionControlItem = GetVersionControlItem();
                        }
                        catch (VersionControlItemNotFoundException)
                        {
                            SetMarginActivated(false);
                            Redraw(false, MarginDrawReason.InternalReason);
                            break;
                        }
                        catch (TeamFoundationServiceUnavailableException)
                        {
                            continue;
                        }

                        if (cancellationToken.IsCancellationRequested)
                            break;

                        if (_versionControlItem == null || versionControlItem.CheckinDate != _versionControlItem.CheckinDate)
                        {
                            _versionControlItem = versionControlItem;
                            DownloadVersionControlItem();
                            Redraw(false, MarginDrawReason.VersionControlItemChanged);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Unhandled exception in ObserveVersionControlItem: " + ex, Properties.Resources.ProductName);
                }
            }
        }

        /// <summary>
        /// Get version control item for current text document.
        /// </summary>
        /// <returns>Committed version of a file in the version control server.</returns>
        /// <exception cref="VersionControlItemNotFoundException">The document file is not linked to version control.</exception>
        /// <exception cref="TeamFoundationServiceUnavailableException">Team Foundation Service unavailable.</exception>
        private Item GetVersionControlItem()
        {
            try
            {
                // Be careful, VersionControlServer.GetItem is slow.
                return _versionControl.GetItem(_textDoc.FilePath, VersionSpec.Latest);
            }
            catch (VersionControlException ex)
            {
                string msg = string.Format("Item not found in repository on path \"{0}\".", _textDoc.FilePath);
                throw new VersionControlItemNotFoundException(msg, ex);
            }
        }

        /// <summary>
        /// Download version control item for current text document to stream.
        /// </summary>
        private void DownloadVersionControlItem()
        {
            _versionControlItemStream = new MemoryStream();
            _versionControlItem.DownloadFile().CopyTo(_versionControlItemStream);
            AppendShiftTokenToStream(_versionControlItemStream, Encoding.GetEncoding(_versionControlItem.Encoding));
        }

        /// <summary>
        /// Update difference between lines asynchronously, if needed, and redraw the margin.
        /// </summary>
        /// <param name="useCache">Use cached differences.</param>
        /// <param name="reason">The reason of redrawing.</param>
        private void Redraw(bool useCache, MarginDrawReason reason)
        {
            try
            {
                if (_textView.IsClosed)
                    return;

                if (!_isActivated)
                {
                    _cachedChangedLines.Clear();
                    RaiseMarginRedraw(reason);
                    return;
                }

                if (useCache)
                {
                    RaiseMarginRedraw(reason);
                    return;
                }

                var task = new Task(() =>
                {
                    lock (_drawLockObject)
                    {
                        try
                        {
                            _cachedChangedLines = GetChangedLineNumbers();
                        }
                        catch (Exception ex)
                        {
                            RaiseExceptionThrown(ex);
                            return;
                        }

                        Redraw(true, reason);
                    }
                });

                task.Start();
            }
            catch (Exception ex)
            {
                RaiseExceptionThrown(ex);
            }
        }

        /// <summary>
        /// Get differences between an original stream and a modified stream.
        /// </summary>
        /// <param name="originalStream">Original stream.</param>
        /// <param name="originalEncoding">Encoding of original stream.</param>
        /// <param name="modifiedStream">Modified stream.</param>
        /// <param name="modifiedEncoding">Encoding of modified stream.</param>
        /// <returns>A summary of the differences between two streams.</returns>
        private DiffSummary GetDifference(Stream originalStream, Encoding originalEncoding, Stream modifiedStream, Encoding modifiedEncoding)
        {
            var diffOptions = new DiffOptions { UseThirdPartyTool = false };

            // TODO: make flag IgnoreLeadingAndTrailingWhiteSpace configurable via "Tools|Options..." dialog (TfsPendingChangesMargin settings).
            diffOptions.Flags = diffOptions.Flags | DiffOptionFlags.IgnoreLeadingAndTrailingWhiteSpace;

            DiffSummary diffSummary;

            lock (_differenceLockObject)
            {
                originalStream.Position = 0;
                modifiedStream.Position = 0;

                diffSummary = DiffUtil.Diff(
                    originalStream,
                    originalEncoding,
                    modifiedStream,
                    modifiedEncoding,
                    diffOptions,
                    true);
            }

            return diffSummary;
        }

        /// <summary>
        /// Get differences between the document's local file and his source control version.
        /// </summary>
        /// <returns>
        /// Collection that contains the result of comparing the document's local file with his source control version.
        /// <para/>Each element is a pair of key and value: the key is a line of text, the value is a type of difference.
        /// </returns>
        private DiffLinesCollection GetChangedLineNumbers()
        {
            Debug.Assert(_textDoc != null, "_textDoc is null.");
            Debug.Assert(_versionControl != null, "_versionControl is null.");

            var textSnapshot = _textView.TextSnapshot;

            Stream sourceStream = new MemoryStream();
            Encoding sourceStreamEncoding = _textDoc.Encoding;
            byte[] textBytes = sourceStreamEncoding.GetBytes(textSnapshot.GetText());
            sourceStream.Write(textBytes, 0, textBytes.Length);
            AppendShiftTokenToStream(sourceStream, sourceStreamEncoding);

            DiffSummary diffSummary = GetDifference(
                _versionControlItemStream,
                Encoding.GetEncoding(_versionControlItem.Encoding),
                sourceStream,
                sourceStreamEncoding);

            var dict = new DiffLinesCollection();
            for (int i = 0; i < diffSummary.Changes.Length; i++)
            {
                IDiffChange diffChange = diffSummary.Changes[i];
                int diffStartLineIndex = diffChange.ModifiedStart;
                int diffEndLineIndex = diffChange.ModifiedEnd - 1;

                DiffChangeType diffType = diffChange.ChangeType;
                if (diffType == DiffChangeType.Change)
                {
                    if (diffChange.OriginalLength >= diffChange.ModifiedLength)
                    {
                        int linesModified = diffChange.ModifiedLength;
                        if (linesModified == 0)
                        {
                            diffType = DiffChangeType.Delete;
                            int linesDeleted = diffChange.OriginalLength - diffChange.ModifiedLength;
                            Debug.Assert(linesDeleted > 0, "linesDeleted must be greater than zero.");
                        }
                    }
                    else
                    {
                        int linesModified = diffChange.OriginalLength;
                        if (linesModified == 0)
                        {
                            diffType = DiffChangeType.Insert;
                            int linesAdded = diffChange.ModifiedLength - diffChange.OriginalLength;
                            Debug.Assert(linesAdded > 0, "linesAdded must be greater than zero.");
                        }
                    }
                }

                if (diffType != DiffChangeType.Delete)
                {
                    for (int k = diffStartLineIndex; k <= diffEndLineIndex; k++)
                        AddLineToDiffLinesCollection(dict, textSnapshot, k, diffType);
                }
                else
                {
                    AddLineToDiffLinesCollection(dict, textSnapshot, diffEndLineIndex, diffType);
                }
            }

            return dict;
        }

        #region Event handlers

        /// <summary>
        /// Event handler that occurs when the <see cref="IWpfTextView.ZoomLevel"/> is set.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event arguments.</param>
        private void OnTextViewZoomLevelChanged(object sender, ZoomLevelChangedEventArgs e)
        {
            Redraw(true, MarginDrawReason.TextViewZoomLevelChanged);
        }

        /// <summary>
        /// Event handler that occurs when the document has been loaded from or saved to disk. 
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event arguments.</param>
        private void OnTextDocFileActionOccurred(object sender, TextDocumentFileActionEventArgs e)
        {
            Redraw(false, MarginDrawReason.TextDocFileActionOccurred);
        }

        /// <summary>
        /// Event handler that occurs when the text editor performs a text line layout.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event arguments.</param>
        private void OnTextViewLayoutChanged(object sender, TextViewLayoutChangedEventArgs e)
        {
            if (AnyTextChanges(e.OldViewState.EditSnapshot.Version, e.NewViewState.EditSnapshot.Version))
                Redraw(false, MarginDrawReason.TextViewTextChanged);
            else if (e.VerticalTranslation || e.NewOrReformattedSpans.Count > 0)
                Redraw(true, MarginDrawReason.TextViewLayoutChanged);
        }

        /// <summary>
        /// Event handler that occurs when <see cref="IEditorFormatMap"/> changes.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event arguments.</param>
        private void OnFormatMapFormatMappingChanged(object sender, FormatItemsEventArgs e)
        {
            _marginSettings.Refresh();
            Redraw(true, MarginDrawReason.EditorFormatMapChanged);
        }

        /// <summary>
        /// Event handler that raised on the commit of a new check-in.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event arguments.</param>
        private void OnVersionControlCommitCheckin(object sender, CommitCheckinEventArgs e)
        {
            if (_versionControlItem == null)
                return;

            string serverItem = _versionControlItem.ServerItem;
            bool itemCommitted = e.Changes.Any(change => change.ServerItem == serverItem);
            if (itemCommitted)
            {
                var task = new Task(() =>
                {
                    lock (_drawLockObject)
                    {
                        try
                        {
                            _versionControlItem = GetVersionControlItem();
                            DownloadVersionControlItem();
                            Redraw(false, MarginDrawReason.VersionControlItemChanged);
                        }
                        catch (VersionControlItemNotFoundException)
                        {
                            SetMarginActivated(false);
                            Redraw(false, MarginDrawReason.InternalReason);
                        }
                        catch (TeamFoundationServiceUnavailableException)
                        {
                        }
                    }
                });

                task.Start();
            }
        }

        /// <summary>
        /// Event handler that occurs when the current project context changes.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event arguments.</param>
        private void OnTfExtProjectContextChanged(object sender, EventArgs e)
        {
            UpdateMargin();
        }

        /// <summary>
        /// Event handler that occurs when the mapping has changed between a character position and its vertical fraction. 
        /// For example, the view may have re-rendered some lines, changing their font size.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event arguments.</param>
        private void OnScrollMapMappingChanged(object sender, EventArgs e)
        {
            Redraw(true, MarginDrawReason.ScrollMapMappingChanged);
        }

        #endregion Event handlers
    }
}