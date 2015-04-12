using System;
using System.Runtime.InteropServices;
using AlekseyNagovitsyn.TfsPendingChangesMargin.Settings;
using EnvDTE;
using Microsoft.VisualStudio.Shell;

namespace AlekseyNagovitsyn.TfsPendingChangesMargin
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </summary>
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid("9ec6e3aa-48cd-4953-b3b7-1a203bfded7f")]
    [ProvideOptionPage(typeof(GeneralSettingsPage), "Tfs Pending Changes Margin", "General", 0, 0, true)]
    public sealed class TfsPendingChangesMarginPackage : Package
    {
        /// <summary>
        /// Returns the settings from the Tools|Options... dialog's Tfs Pending Changes Margin|General section
        /// </summary>
        public static EnvDTE.Properties GetGeneralSettings()
        {
            var dte = (DTE)GetGlobalService(typeof(DTE));
            return dte.Properties["Tfs Pending Changes Margin", "General"];
        }

        /// <summary>
        /// Event raised when the settings on the <see cref="GeneralSettingsPage"/> are changed
        /// </summary>
        public static event Action GeneralSettingsChanged;

        /// <summary>
        /// Raises the <see cref="GeneralSettingsChanged"/> event.
        /// </summary>
        public static void RaiseGeneralSettingsChanged()
        {
            if (GeneralSettingsChanged != null)
                GeneralSettingsChanged();
        }
    }
}
