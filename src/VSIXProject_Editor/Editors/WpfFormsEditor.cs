using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using VSExtensibilityHelper.Core.Base;
using VSExtensibilityHelper.Core.Service;

namespace VSIXProject_Editor.Editors
{
    [Guid("08CD1411-7D9D-4CF0-8897-BDDA012F7AE1")]
    public sealed class WpfFormsEditorFactory : BaseEditorFactory<WpfEditorEditorPane>
    {
    }

    public class WpfEditorEditorPane : BaseWpfEditorPane<WpfFormsEditorFactory, Controls.Wpf.WpfUserControl>
    {
        #region Methods

        protected override string GetFileExtension()
        {
            return ".xam";
        }

        protected override void LoadFile(string fileName)
        {
            PaneService.Initialize(ServiceLocator.GetInstance<IServiceProvider>(), "My Pane");
            PaneService.Log($"Loading file: {fileName}");
            PaneService.Activate();
        }

        protected override void SaveFile(string fileName)
        {

        }

        #endregion Methods
    }
}
