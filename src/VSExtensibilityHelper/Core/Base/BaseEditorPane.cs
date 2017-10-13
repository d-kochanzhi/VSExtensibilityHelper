/*
 https://github.com/d-kochanzhi/VSExtensibilityHelper
 http://dzsoft.ru/post/VSExtensibilityHelper 
 */
using System;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using VSExtensibilityHelper.Core.Service;

namespace VSExtensibilityHelper.Core.Base
{
    /// <summary>
    /// Base class for editor pane
    /// </summary>
    /// <typeparam name="TFactory"></typeparam>
    public abstract class BaseEditorPane<TFactory> :
         WindowPane,
            IPersistFileFormat,
            IVsPersistDocData
      where TFactory : IVsEditorFactory
    {
        #region Fields

        private const char EndLineChar = (char)10;

        private const uint FileFormatIndex = 0;

        private readonly string _FileExtensionUsed;

        private string _FileName;

        private bool _GettingCheckoutStatus;

        private bool _IsDirty;

        private bool _Loading;

        private bool _NoScribbleMode;

        #endregion Fields

        #region Constructors

        protected BaseEditorPane() : base(null)
        {
            _FileExtensionUsed = GetFileExtension();

            PrivateInit();
        }

        #endregion Constructors

        #region Properties

        public Guid FactoryGuid
        {
            get { return typeof(TFactory).GUID; }
        }

        public string FileExtensionUsed
        {
            get { return _FileExtensionUsed; }
        }

        public bool IsDirty
        {
            get { return _IsDirty; }
        }

        #endregion Properties

        #region Methods

        private bool CanEditFile()
        {
            if (_GettingCheckoutStatus)
            {
                return false;
            }

            try
            {
                _GettingCheckoutStatus = true;

                IVsQueryEditQuerySave2 queryEditQuerySave = (IVsQueryEditQuerySave2)GetService(typeof(SVsQueryEditQuerySave));

                string[] documents = { _FileName };
                uint result;
                uint outFlags;

                int hr = queryEditQuerySave.QueryEditFiles(
                  0,
                  1,
                  documents,
                  null,
                  null,
                  out result,
                  out outFlags
                  );
                if (ErrorHandler.Succeeded(hr) && (result == (uint)tagVSQueryEditResult.QER_EditOK))
                {
                    return true;
                }
            }
            finally
            {
                _GettingCheckoutStatus = false;
            }
            return false;
        }

        private void NotifyDocChanged()
        {
            IVsHierarchy hierarchy;
            uint itemID;
            IntPtr docData;
            uint docCookie = 0;

            if (_FileName.Length == 0)
                return;

            IVsRunningDocumentTable runningDocTable = (IVsRunningDocumentTable)GetService(
                typeof(SVsRunningDocumentTable));

            try
            {
                ErrorHandler.ThrowOnFailure(runningDocTable.FindAndLockDocument((uint)_VSRDTFLAGS.RDT_ReadLock,
                    _FileName, out hierarchy, out itemID, out docData, out docCookie));

                ErrorHandler.ThrowOnFailure(runningDocTable.NotifyDocumentChanged(docCookie,
                    (uint)__VSRDTATTRIB.RDTA_DocDataReloaded));
            }
            finally
            {
                if (runningDocTable != null)
                    ErrorHandler.ThrowOnFailure(runningDocTable.UnlockDocument((uint)_VSRDTFLAGS.RDT_ReadLock,
                        docCookie));
            }
        }

        private void PrivateInit()
        {
            _NoScribbleMode = false;
            _Loading = false;
            _GettingCheckoutStatus = false;
        }

        protected abstract string GetFileExtension();

        protected abstract void LoadFile(string fileName);

        protected virtual int OnCloseEditor()
        {
            return VSConstants.S_OK;
        }

        protected virtual void OnContentChanged()
        {
            if (!_Loading)
            {
                if (!_IsDirty)
                {
                    if (!CanEditFile())
                    {
                        return;
                    }

                    _IsDirty = true;
                }
            }
        }

        protected virtual int OnGetCurFile(out string ppszFilename, out uint pnFormatIndex)
        {
            pnFormatIndex = FileFormatIndex;
            ppszFilename = _FileName;
            return VSConstants.S_OK;
        }

        protected virtual int OnGetFormatList(out string ppszFormatList)
        {
            string formatList =
              string.Format(CultureInfo.CurrentCulture,
              "Editor Files (*{0}){1}*{0}{1}{1}",
              FileExtensionUsed, EndLineChar);
            ppszFormatList = formatList;
            return VSConstants.S_OK;
        }

        protected virtual int OnInitNew(uint nFormatIndex)
        {
            if (nFormatIndex != FileFormatIndex)
            {
                throw new ArgumentException("Unknown format");
            }
            _IsDirty = false;
            return VSConstants.S_OK;
        }

        protected virtual int OnIsDirty(out int pfIsDirty)
        {
            pfIsDirty = _IsDirty ? 1 : 0;
            return VSConstants.S_OK;
        }

        protected virtual int OnSaveCompleted(string pszFilename)
        {
            return _NoScribbleMode ? VSConstants.S_FALSE : VSConstants.S_OK;
        }

        protected abstract void SaveFile(string fileName);

        int IVsPersistDocData.Close()
        {
            return OnCloseEditor();
        }

        int IPersist.GetClassID(out Guid pClassID)
        {
            pClassID = FactoryGuid;
            return VSConstants.S_OK;
        }

        int IPersistFileFormat.GetClassID(out Guid pClassID)
        {
            pClassID = FactoryGuid;
            return VSConstants.S_OK;
        }

        int IPersistFileFormat.GetCurFile(out string ppszFilename, out uint pnFormatIndex)
        {
            return OnGetCurFile(out ppszFilename, out pnFormatIndex);
        }

        int IPersistFileFormat.GetFormatList(out string ppszFormatList)
        {
            return OnGetFormatList(out ppszFormatList);
        }

        int IVsPersistDocData.GetGuidEditorType(out Guid pClassID)
        {
            pClassID = FactoryGuid;
            return VSConstants.S_OK;
        }

        int IPersistFileFormat.InitNew(uint nFormatIndex)
        {
            return OnInitNew(nFormatIndex);
        }

        int IPersistFileFormat.IsDirty(out int pfIsDirty)
        {
            return OnIsDirty(out pfIsDirty);
        }

        int IVsPersistDocData.IsDocDataDirty(out int pfDirty)
        {
            return ((IPersistFileFormat)this).IsDirty(out pfDirty);
        }

        int IVsPersistDocData.IsDocDataReloadable(out int pfReloadable)
        {
            pfReloadable = 1;
            return VSConstants.S_OK;
        }

        int IPersistFileFormat.Load(string pszFilename, uint grfMode, int fReadOnly)
        {
            if ((pszFilename == null) && ((_FileName == null) || (_FileName.Length == 0)))
                throw new ArgumentNullException("pszFilename");

            _Loading = true;
            int hr = VSConstants.S_OK;
            try
            {
                bool isReload = false;
                if (pszFilename == null)
                {
                    isReload = true;
                }

                ServiceLocator.GetGlobalService<SVsUIShell, IVsUIShell>().SetWaitCursor();

                if (!isReload)
                {
                    _FileName = pszFilename;
                }
                LoadFile(_FileName);
                _IsDirty = false;

                NotifyDocChanged();
            }
            finally
            {
                _Loading = false;
            }
            return hr;
        }

        int IVsPersistDocData.LoadDocData(string pszMkDocument)
        {
            return ((IPersistFileFormat)this).Load(pszMkDocument, 0, 0);
        }

        int IVsPersistDocData.OnRegisterDocData(uint docCookie, IVsHierarchy pHierNew,
          uint itemidNew)
        {
            return VSConstants.S_OK;
        }

        int IVsPersistDocData.ReloadDocData(uint grfFlags)
        {
            return ((IPersistFileFormat)this).Load(null, grfFlags, 0);
        }

        int IVsPersistDocData.RenameDocData(uint grfAttribs, IVsHierarchy pHierNew,
          uint itemidNew, string pszMkDocumentNew)
        {
            return VSConstants.S_OK;
        }

        int IPersistFileFormat.Save(string pszFilename, int fRemember, uint nFormatIndex)
        {
            _NoScribbleMode = true;
            try
            {
                if (pszFilename == null || pszFilename == _FileName)
                {
                    SaveFile(_FileName);
                    _IsDirty = false;
                }
                else
                {
                    if (fRemember != 0)
                    {
                        _FileName = pszFilename;
                        SaveFile(_FileName);
                        _IsDirty = false;
                    }
                    else
                    {
                        SaveFile(pszFilename);
                    }
                }
            }
            finally
            {
                _NoScribbleMode = false;
            }
            return VSConstants.S_OK;
        }

        int IPersistFileFormat.SaveCompleted(string pszFilename)
        {
            return OnSaveCompleted(pszFilename);
        }

        [SuppressMessage("Microsoft.Globalization", "CA1303:DoNotPassLiteralsAsLocalizedParameters")]
        int IVsPersistDocData.SaveDocData(VSSAVEFLAGS dwSave, out string pbstrMkDocumentNew, out int pfSaveCanceled)
        {
            pbstrMkDocumentNew = null;
            pfSaveCanceled = 0;
            int hr;

            switch (dwSave)
            {
                case VSSAVEFLAGS.VSSAVE_Save:
                case VSSAVEFLAGS.VSSAVE_SilentSave:
                    {
                        IVsQueryEditQuerySave2 queryEditQuerySave =
                          (IVsQueryEditQuerySave2)GetService(typeof(SVsQueryEditQuerySave));

                        uint result;
                        hr = queryEditQuerySave.QuerySaveFile(
                          _FileName,
                          0,
                          null,
                          out result);

                        if (ErrorHandler.Failed(hr))
                        {
                            return hr;
                        }

                        switch ((tagVSQuerySaveResult)result)
                        {
                            case tagVSQuerySaveResult.QSR_NoSave_Cancel:
                                pfSaveCanceled = ~0;
                                break;

                            case tagVSQuerySaveResult.QSR_SaveOK:
                                {
                                    hr = ServiceLocator.GetGlobalService<SVsUIShell, IVsUIShell>().SaveDocDataToFile(
                                        dwSave, this, _FileName, out pbstrMkDocumentNew, out pfSaveCanceled);
                                    if (ErrorHandler.Failed(hr))
                                        return hr;
                                }
                                break;

                            case tagVSQuerySaveResult.QSR_ForceSaveAs:
                                {
                                    hr = ServiceLocator.GetGlobalService<SVsUIShell, IVsUIShell>().SaveDocDataToFile(
                                        VSSAVEFLAGS.VSSAVE_SaveAs, this, _FileName, out pbstrMkDocumentNew,
                                        out pfSaveCanceled);
                                    if (ErrorHandler.Failed(hr))
                                        return hr;
                                }
                                break;

                            case tagVSQuerySaveResult.QSR_NoSave_Continue:
                                break;

                            default:
                                throw new InvalidOperationException("Unsupported result from Query Edit/Query Save");
                        }
                        break;
                    }
                case VSSAVEFLAGS.VSSAVE_SaveAs:
                case VSSAVEFLAGS.VSSAVE_SaveCopyAs:
                    {
                        if (String.Compare(FileExtensionUsed, Path.GetExtension(_FileName), true, CultureInfo.CurrentCulture) != 0)
                        {
                            _FileName += FileExtensionUsed;
                        }
                        hr = ServiceLocator.GetGlobalService<SVsUIShell, IVsUIShell>().SaveDocDataToFile(dwSave,
                            this, _FileName, out pbstrMkDocumentNew, out pfSaveCanceled);
                        if (ErrorHandler.Failed(hr))
                            return hr;
                        break;
                    }
                default:
                    throw new ArgumentException("Unsupported save flag value");
            }
            return VSConstants.S_OK;
        }

        int IVsPersistDocData.SetUntitledDocPath(string pszDocDataPath)
        {
            return ((IPersistFileFormat)this).InitNew(FileFormatIndex);
        }

        #endregion Methods
    }
}