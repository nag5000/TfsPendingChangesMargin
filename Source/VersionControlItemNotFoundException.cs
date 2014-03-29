using System;

namespace AlekseyNagovitsyn.TfsPendingChangesMargin
{
    /// <summary>
    /// The exception that is thrown when an attempt to access an item that does not exist in version control fails.
    /// </summary>
    internal class VersionControlItemNotFoundException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="VersionControlItemNotFoundException"/> class.
        /// </summary>
        public VersionControlItemNotFoundException()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="VersionControlItemNotFoundException"/> class 
        /// with a specified error message.
        /// </summary>
        /// <param name="message">A description of the error.</param>
        public VersionControlItemNotFoundException(string message) 
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="VersionControlItemNotFoundException"/> class 
        /// with a specified error message and a reference to the inner exception that is the cause of this exception.
        /// </summary>
        /// <param name="message">A description of the error.</param>
        /// <param name="innerException">The exception that is the cause of the current exception.</param>
        public VersionControlItemNotFoundException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}