/*
 https://github.com/d-kochanzhi/VSExtensibilityHelper
 http://dzsoft.ru/post/VSExtensibilityHelper 
 */
using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using IOleServiceProvider = Microsoft.VisualStudio.OLE.Interop.IServiceProvider;

namespace VSExtensibilityHelper.Core.Base
{
    /// <summary>
    /// Base class for factory
    /// </summary>
    /// <typeparam name="TEditorPane"></typeparam>
    public class BaseEditorFactory<TEditorPane> :
      IVsEditorFactory,
      IDisposable
      where TEditorPane :
        WindowPane, IOleCommandTarget, IVsPersistDocData, IPersistFileFormat, new()
    {
        #region Fields

        private ServiceProvider _ServiceProvider;

        #endregion Fields

        #region Constructors

        public BaseEditorFactory()
        {
            Trace.WriteLine(
              string.Format(CultureInfo.CurrentCulture,
              "Entering {0} constructor", typeof(TEditorPane)));
        }

        #endregion Constructors

        #region Methods

        private void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_ServiceProvider != null)
                {
                    _ServiceProvider.Dispose();
                    _ServiceProvider = null;
                }
            }
        }

        public virtual int Close()
        {
            return VSConstants.S_OK;
        }

        [EnvironmentPermission(SecurityAction.Demand, Unrestricted = true)]
        public virtual int CreateEditorInstance(
          uint grfCreateDoc,
          string pszMkDocument,
          string pszPhysicalView,
          IVsHierarchy pvHier,
          uint itemid,
          IntPtr punkDocDataExisting,
          out IntPtr ppunkDocView,
          out IntPtr ppunkDocData,
          out string pbstrEditorCaption,
          out Guid pguidCmdUI,
          out int pgrfCDW)
        {
            Trace.WriteLine(string.Format(CultureInfo.CurrentCulture,
              "Entering {0} CreateEditorInstance()", ToString()));

            ppunkDocView = IntPtr.Zero;
            ppunkDocData = IntPtr.Zero;
            pguidCmdUI = GetType().GUID;
            pgrfCDW = 0;
            pbstrEditorCaption = null;

            if ((grfCreateDoc & (VSConstants.CEF_OPENFILE | VSConstants.CEF_SILENT)) == 0)
            {
                return VSConstants.E_INVALIDARG;
            }
            if (punkDocDataExisting != IntPtr.Zero)
            {
                return VSConstants.VS_E_INCOMPATIBLEDOCDATA;
            }

            TEditorPane newEditor = new TEditorPane();
            ppunkDocView = Marshal.GetIUnknownForObject(newEditor);
            ppunkDocData = Marshal.GetIUnknownForObject(newEditor);
            pbstrEditorCaption = "";

            return VSConstants.S_OK;
        }

        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        public object GetService(Type serviceType)
        {
            return _ServiceProvider.GetService(serviceType);
        }

        public virtual int MapLogicalView(ref Guid logicalView, out string physicalView)
        {
            physicalView = null;

            if (VSConstants.LOGVIEWID_Primary == logicalView)
            {
                return VSConstants.S_OK;
            }
            else
            {
                return VSConstants.E_NOTIMPL;
            }
        }

        public virtual int SetSite(IOleServiceProvider serviceProvider)
        {
            _ServiceProvider = new ServiceProvider(serviceProvider);
            return VSConstants.S_OK;
        }

        #endregion Methods
    }
}