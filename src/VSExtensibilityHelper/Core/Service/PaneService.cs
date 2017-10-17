using System;
using Microsoft.VisualStudio.Shell.Interop;

namespace VSExtensibilityHelper.Core.Service
{
    /*
    based on https://raw.githubusercontent.com/madskristensen/VsixLogger/master/src/Logger.cs
    */

    public static class PaneService
    {
        #region Fields

        private static Guid? _guid;
        private static string _name;
        private static IVsOutputWindowPane _pane;
        private static IServiceProvider _provider;
        private static object _syncRoot = new object();

        #endregion Fields

        #region Methods

        /// <summary>
        /// Get a pane. 
        /// You can use  Microsoft.VisualStudio.VSConstants.OutputWindowPaneGuid. 
        /// </summary>
        /// <returns></returns>
        private static bool EnsurePane()
        {
            if (_pane == null)
            {
                
                lock (_syncRoot)
                {
                    if (_pane == null)
                    {
                        if (!_guid.HasValue)
                            _guid = Guid.NewGuid();

                        Guid _temp = _guid.Value;
                        IVsOutputWindow output = (IVsOutputWindow)_provider.GetService(typeof(SVsOutputWindow));

                        output.GetPane(_guid.Value, out _pane);

                        if (_pane == null)
                        {
                            output.CreatePane(
                                ref _temp,
                                string.IsNullOrEmpty(_name)? "VSExtensibilityHelper - Pane" : _name,
                                1,
                                1);

                            output.GetPane(ref _temp, out _pane);
                        }
                    }
                }
            }

            return _pane != null;
        }

        public static void Activate()
        {
            if (_pane != null)
            {
                _pane.Activate();
            }
        }

        public static void Clear()
        {
            if (_pane != null)
            {
                _pane.Clear();
            }
        }

        public static void DeletePane()
        {
            if (_pane != null)
            {
                try
                {
                    IVsOutputWindow output = (IVsOutputWindow)_provider.GetService(typeof(SVsOutputWindow));
                    Guid _temp = _guid.Value;
                    output.DeletePane(ref _temp);
                    _guid = null;
                    _pane = null;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.Write(ex);
                }
            }
        }

        public static void Initialize(IServiceProvider provider, string name)
        {
            _pane = null;
            _provider = provider;
            _name = name;
            _guid = null;
        }

        public static void Initialize(IServiceProvider provider, Guid paneGuid)
        {
            _pane = null;
            _provider = provider;
            _guid = paneGuid;
            _name = null;
        }

        public static void Log(string message)
        {
            if (string.IsNullOrEmpty(message))
                return;

            try
            {
                if (EnsurePane())
                {
                    _pane.OutputStringThreadSafe($"[{DateTime.Now.ToShortTimeString()}] : {message}" + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Write(ex);
            }
        }

        public static void Log(Exception ex)
        {
            if (ex != null)
            {
                Log(ex.ToString());
            }
        }

        #endregion Methods
    }
}