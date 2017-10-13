using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using VSExtensibilityHelper.Core.Base;

namespace VSIXProject_Editor.Editors
{
    [Guid("21411E52-A07A-4304-A162-8C87FD7056EA")]
    public sealed class WinFormsEditorFactory : BaseEditorFactory<WinFormsEditorEditorPane>
    {
    }

    public class WinFormsEditorEditorPane : BaseWinFormsEditorPane<WinFormsEditorFactory, Controls.WinForms.WinFormUserControl>
    {
        #region Methods

        protected override string GetFileExtension()
        {
            return ".win";
        }

        protected override void LoadFile(string fileName)
        {
           
        }

        protected override void SaveFile(string fileName)
        {
           
        }     

        #endregion Methods
    }
}
