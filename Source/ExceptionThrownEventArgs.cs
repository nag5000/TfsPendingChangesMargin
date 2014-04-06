using System;

namespace AlekseyNagovitsyn.TfsPendingChangesMargin
{
    /// <summary>
    /// Arguments for event that raised when an unhandled exception was catched.
    /// </summary>
    internal class ExceptionThrownEventArgs : EventArgs
    {
        /// <summary>
        /// The occurred exception.
        /// </summary>
        private readonly Exception _exception;

        /// <summary>
        /// Gets the occurred exception.
        /// </summary>
        public Exception Exception
        {
            get { return _exception; }
        }

        /// <summary>
        /// Gets or sets that the exception was handled.
        /// </summary>
        public bool Handled { get; set; }

        /// <summary>
        /// Initialize a new instance of the <see cref="ExceptionThrownEventArgs"/> class.
        /// </summary>
        /// <param name="exception">The occurred exception.</param>
        public ExceptionThrownEventArgs(Exception exception)
        {
            _exception = exception;
        }
    }
}