﻿using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using VsServiceProvider = Microsoft.VisualStudio.OLE.Interop.IServiceProvider;

namespace VSExtensibilityHelper.Core.Service
{
    /// <summary>
    /// This class unifies all the different ways of getting services within visual studio.
    /// </summary>
    public static class ServiceLocator
    {
        #region Properties

        public static IServiceProvider PackageServiceProvider
        {
            get;
            private set;
        }

        #endregion Properties

        #region Methods

        private static TService GetComponentModelService<TService>() where TService : class
        {
            IComponentModel componentModel = GetGlobalService<SComponentModel, IComponentModel>();
            return componentModel.GetService<TService>();
        }

        private static TService GetDTEService<TService>() where TService : class
        {
            var dte = GetGlobalService<SDTE, DTE>();
            return (TService)QueryService(dte, typeof(TService));
        }

        private static IServiceProvider GetServiceProvider()
        {
            var dte = GetGlobalService<SDTE, DTE>();
            return GetServiceProvider(dte);
        }

        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "The caller is responsible for disposing this")]
        private static IServiceProvider GetServiceProvider(_DTE dte)
        {
            IServiceProvider serviceProvider = new ServiceProvider(dte as VsServiceProvider);
            Debug.Assert(serviceProvider != null, "Service provider is null");
            return serviceProvider;
        }

        private static object QueryService(_DTE dte, Type serviceType)
        {
            Guid guidService = serviceType.GUID;
            Guid riid = guidService;
            var serviceProvider = dte as VsServiceProvider;

            IntPtr servicePtr;
            int hr = serviceProvider.QueryService(ref guidService, ref riid, out servicePtr);

            if (hr != VSConstants.S_OK)
            {
                // We didn't find the service so return null
                return null;
            }

            object service = null;

            if (servicePtr != IntPtr.Zero)
            {
                service = Marshal.GetObjectForIUnknown(servicePtr);
                Marshal.Release(servicePtr);
            }

            return service;
        }

        public static TInterface GetGlobalService<TService, TInterface>() where TInterface : class
        {
            if (PackageServiceProvider != null)
            {
                TInterface service = PackageServiceProvider.GetService(typeof(TService)) as TInterface;
                if (service != null)
                {
                    return service;
                }
            }

            return (TInterface)Package.GetGlobalService(typeof(TService));
        }

        public static TService GetInstance<TService>() where TService : class
        {
            // Special case IServiceProvider
            if (typeof(TService) == typeof(IServiceProvider))
            {
                return (TService)GetServiceProvider();
            }

            // then try to find the service as a component model, then try dte then lastly try global service
            // avoid calling GetGlobalService() from within the Initialize() method of NuGetPackage class.
            // Doing so is illegal and may cause VS to hang. As a result of that, we defer calling GetGlobalService to the last option.
            return GetDTEService<TService>() ??
                   GetComponentModelService<TService>() ??
                   GetGlobalService<TService, TService>();
        }

        public static void InitializePackageServiceProvider(IServiceProvider provider)
        {
            if (provider == null)
            {
                throw new ArgumentNullException("provider");
            }

            PackageServiceProvider = provider;
        }

        #endregion Methods
    }
}