using System;
using System.Runtime.InteropServices;
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
    [ProvideOptionPage(typeof(Settings.OptionPage), "Tfs Pending Changes Margin", "General", 0, 0, true)]
    public sealed class TfsPendingChangesMarginPackage : Package
    {
        public static bool IgnoreLeadingAndTrailingWhiteSpace
        {
            get
            {
                var dte = (DTE)GetGlobalService(typeof(DTE));
                EnvDTE.Properties properties = dte.Properties["Tfs Pending Changes Margin", "General"];
                return (bool)properties.Item("IgnoreLeadingAndTrailingWhiteSpace").Value;
            }
        }
    }
}
