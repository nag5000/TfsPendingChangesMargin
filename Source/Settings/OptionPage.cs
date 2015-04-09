using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;

namespace AlekseyNagovitsyn.TfsPendingChangesMargin.Settings
{
    /// <summary>
    /// This page is a basic grid including any settings used by the TfsPendingChangesMargin extension.
    /// 
    /// The ProvideOptionPageAttribute on the TfsPendingChangesMarginPackage is used to add this page
    /// to the Visual Studio Tools-> Options dialog.
    /// </summary>
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [CLSCompliant(false), ComVisible(true)]
    public class OptionPage : DialogPage
    {
        [Category("Diff Options")]
        [DisplayName("Ignore leading and trailing white space")]
        [Description("If True, the margin will not show lines where the only changes are to leading or trailing white space")]
        public bool IgnoreLeadingAndTrailingWhiteSpace { get; set; }

        public OptionPage()
        {
            // TODO: in C# 6, default values for properties can be specified inline so this constructor can be removed
            IgnoreLeadingAndTrailingWhiteSpace = true;
        }
    }
}
