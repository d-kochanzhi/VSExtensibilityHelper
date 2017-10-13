using System;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace VSExtensibilityHelper.Core.Service
{
    public static class MsgBoxService
    {
        #region Methods

        public static void Error(Exception ex, string title = "Error")
        {
            VsShellUtilities.ShowMessageBox(
                    ServiceLocator.GetInstance<IServiceProvider>(),
                    $"{ex.Message}\n{ex.StackTrace}",
                    title,
                    OLEMSGICON.OLEMSGICON_CRITICAL,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        public static void Info(string message, string title = "Info")
        {
            VsShellUtilities.ShowMessageBox(
                ServiceLocator.GetInstance<IServiceProvider>(),
                message,
                title,
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        #endregion Methods
    }
}