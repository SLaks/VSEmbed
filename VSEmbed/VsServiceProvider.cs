using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.ComponentModel.Composition.Primitives;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using VSEmbed.Services;
using OLE = Microsoft.VisualStudio.OLE.Interop;
using Shell = Microsoft.VisualStudio.Shell;

namespace VSEmbed {
	///<summary>An out-of-process implementation of Visual Studio's singleton OLE ServiceProvider.</summary>
	///<remarks>
	/// Visual Studio services use this class, both through MEF SVsServiceProvider and
	/// Shell.ServiceProvider.GlobalProvider, to load COM services.  I need to provide
	/// every service used in editor and theme codepaths.  Most of the service methods
	/// are never actually called.
	/// This must be initialized before theme dictionaries or editor services are used
	///</remarks>
	public class VsServiceProvider : OLE.IServiceProvider, SVsServiceProvider {
		public static VsServiceProvider Instance { get; private set; }

		///<summary>Creates the global service provider and populates it with the services we need.</summary>
		[SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "These objects become global and must not be disposed yet")]
		public static void Initialize() {
			// If we're in a real VS process, or if we already initialized, do nothing.
			if (Shell.ServiceProvider.GlobalProvider.GetService(typeof(SVsSettingsManager)) != null)
				return;

			// If the App() ctor didn't set a version, we're in the designer
			// The designer pre-loads our referenced assemblies, so we can't
			// use any other version.
			if (VsLoader.VsVersion == null)
				VsLoader.Load(new Version(11, 0, 0, 0));

			var esm = ExternalSettingsManager.CreateForApplication(Path.Combine(VsLoader.InstallationDirectory, "devenv.exe"));
			var sp = new VsServiceProvider {
				UIShell = new ThemedVsUIShell(),
				serviceInstances = {
					// Used by Shell.ServiceProvider initialization
					{ typeof(SVsActivityLog).GUID, new StubVsActivityLog() },

					// Used by ColorThemeService
					{ typeof(SVsSettingsManager).GUID, new SettingsManagerWrapper(esm) },

					// Used by Shell.VsResourceKeys
					{ new Guid("45652379-D0E3-4EA0-8B60-F2579AA29C93"), new SimpleVsWindowManager() },

					// Used by KnownUIContexts
					{ typeof(IVsMonitorSelection).GUID, new StubVsMonitorSelection() },

					// Used by ShimCodeLensPresenterStyle
					{ typeof(SUIHostLocale).GUID, new SystemUIHostLocale() },
					{ typeof(SVsFontAndColorCacheManager).GUID, new StubVsFontAndColorCacheManager() },

					// Used by Roslyn's VisualStudioWaitIndicator
					{ typeof(SVsThreadedWaitDialogFactory).GUID, new BaseWaitDialogFactory() },

					// Used by Dev14's VsImageLoader, which is needed for Roslyn IntelliSense
					{ typeof(SVsAppId).GUID, new SimpleVsAppId() },

					// Used by DocumentPeekResult; service is SVsUIThreadInvokerPrivate
					{ new Guid("72FD1033-2341-4249-8113-EF67745D84EA"), new AppDispatcherInvoker() },
				}
			};

			sp.AddService(typeof(SVsUIShell), sp.UIShell);

			Shell.ServiceProvider.CreateFromSetSite(sp);
			Instance = sp;

			// The designer loads Microsoft.VisualStudio.Shell.XX.0,
			// which we cannot reference directly (to avoid breaking
			// older versions). Therefore, I set the global property
			// for every available version using Reflection instead.
			foreach (var vsVersion in VsLoader.FindAllVersions().Where(v => v.Major >= 10)) {
				var type = Type.GetType("Microsoft.VisualStudio.Shell.ServiceProvider, Microsoft.VisualStudio.Shell." + vsVersion.ToString(2));
				if (type == null)
					continue;
				type.GetMethod("CreateFromSetSite", BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Static)
					.Invoke(null, new[] { sp });
			}
		}

		///<summary>Gets the MEF IComponentModel installed in this ServiceProvider, if any.</summary>
		public IComponentModel ComponentModel { get; private set; }

		///<summary>Registers a MEF container in this ServiceProvider.</summary>
		///<remarks>
		/// This is used to provide the IComponentModel service, which is used by many parts of Roslyn and the editor.
		/// It's also used by our TextViewHost wrapper control to access the MEF container.
		///</remarks>
		public void SetMefContainer(IComponentModel container) {
			ComponentModel = container;
			AddService(typeof(SComponentModel), ComponentModel);
		}

		public ThemedVsUIShell UIShell { get; private set; }

		readonly Dictionary<Guid, object> serviceInstances = new Dictionary<Guid, object>();

		///<summary>Adds a new service to the provider (or replaces an existing service).</summary>
		public void AddService(Type serviceType, object instance) { AddService(serviceType.GUID, instance); }
		///<summary>Adds a new service to the provider (or replaces an existing service).</summary>
		public void AddService(Guid serviceGuid, object instance) {
			serviceInstances[serviceGuid] = instance;
		}

		int OLE.IServiceProvider.QueryService([ComAliasName("Microsoft.VisualStudio.OLE.Interop.REFGUID")]ref Guid guidService, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.REFIID")]ref Guid riid, out IntPtr ppvObject) {
			object result;
			if (!serviceInstances.TryGetValue(guidService, out result)) {
				ppvObject = IntPtr.Zero;
				return VSConstants.E_NOINTERFACE;
			}
			if (riid == VSConstants.IID_IUnknown) {
				ppvObject = Marshal.GetIUnknownForObject(result);
				return VSConstants.S_OK;
			}

			IntPtr unk = IntPtr.Zero;
			try {
				unk = Marshal.GetIUnknownForObject(result);
				result = Marshal.QueryInterface(unk, ref riid, out ppvObject);
			} finally {
				if (unk != IntPtr.Zero)
					Marshal.Release(unk);
			}
			return VSConstants.S_OK;
		}

		public object GetService(Type serviceType) {
			object result;
			serviceInstances.TryGetValue(serviceType.GUID, out result);
			return result;
		}
	}
}
