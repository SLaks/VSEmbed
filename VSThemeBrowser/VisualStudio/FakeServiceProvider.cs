using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using Microsoft.Internal.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using ServiceProviderRegistration = Microsoft.VisualStudio.Shell.ServiceProvider;

namespace VSThemeBrowser.VisualStudio {
	class FakeServiceProvider : Microsoft.VisualStudio.OLE.Interop.IServiceProvider {
		[SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "These objects become global and must not be disposed yet")]
		public static void Initialize() {
			if (ServiceProviderRegistration.GlobalProvider.GetService(typeof(SVsSettingsManager)) != null)
				return;

			if (VsLoader.VsVersion==null) {		// If the App() ctor didn't set this, we're in the designer
				VsLoader.Initialize(new Version(12, 0, 0, 0));
			}

			var esm = ExternalSettingsManager.CreateForApplication(Path.Combine(VsLoader.GetVersionPath(VsLoader.VsVersion), "devenv.exe"), "Exp");	// FindVsVersions().LastOrDefault().ToString()));
			var sp = new FakeServiceProvider {
				serviceInstances =
				{
					// Used by ServiceProvider
					{ typeof(SVsActivityLog).GUID, new DummyLog() },
					{ typeof(SVsSettingsManager).GUID, new SettingsWrapper(esm) },
					{ new Guid("45652379-D0E3-4EA0-8B60-F2579AA29C93"), new DummyWindowManager() },
					// Activator.CreateInstance(Type.GetType("Microsoft.VisualStudio.Platform.WindowManagement.WindowManagerService, Microsoft.VisualStudio.Platform.WindowManagement, Version=14.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"))

					// Used by KnownUIContexts
					{ typeof(IVsMonitorSelection).GUID, new DummyVsMonitorSelection() },
				}
			};

			ServiceProviderRegistration.CreateFromSetSite(sp);

			// The designer loads Microsoft.VisualStudio.Shell.XX.0,
			// which we cannot reference directly (to avoid breaking
			// older versions). Therefore, I set the global property
			// for every available version using Reflection instead.
			foreach (var vsVersion in VsLoader.FindVsVersions().Where(v => v.Major >= 10)) {
				var type = Type.GetType("Microsoft.VisualStudio.Shell.ServiceProvider, Microsoft.VisualStudio.Shell." + vsVersion.ToString(2));
				if (type == null)
					continue;
				type.GetMethod("CreateFromSetSite", BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Static)
					.Invoke(null, new[] { sp });
			}
		}

		readonly Dictionary<Guid, object> serviceInstances = new Dictionary<Guid, object>();

		public int QueryService([ComAliasName("Microsoft.VisualStudio.OLE.Interop.REFGUID")]ref Guid guidService, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.REFIID")]ref Guid riid, out IntPtr ppvObject) {
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

		class SettingsWrapper : IVsSettingsManager, IDisposable {
			readonly ExternalSettingsManager inner;

			public SettingsWrapper(ExternalSettingsManager inner) {
				this.inner = inner;
			}

			public void Dispose() {
				inner.Dispose();
			}

			public int GetApplicationDataFolder([ComAliasName("Microsoft.VisualStudio.Shell.Interop.VSAPPLICATIONDATAFOLDER")]uint folder, out string folderPath) {
				folderPath = inner.GetApplicationDataFolder((ApplicationDataFolder)folder);
				return 0;
			}

			public int GetCollectionScopes([ComAliasName("OLE.LPCOLESTR")]string collectionPath, [ComAliasName("Microsoft.VisualStudio.Shell.Interop.VSENCLOSINGSCOPES")]out uint scopes) {
				scopes = (uint)inner.GetCollectionScopes(collectionPath);
				return 0;
			}

			public int GetCommonExtensionsSearchPaths([ComAliasName("OLE.ULONG")]uint paths, string[] commonExtensionsPaths, [ComAliasName("OLE.ULONG")]out uint actualPaths) {
				if (commonExtensionsPaths == null)
					actualPaths = (uint)inner.GetCommonExtensionsSearchPaths().Count();
				else {
					inner.GetCommonExtensionsSearchPaths().ToList().CopyTo(commonExtensionsPaths);
					actualPaths = (uint)commonExtensionsPaths.Length;
				}

				return 0;
			}

			public int GetPropertyScopes([ComAliasName("OLE.LPCOLESTR")]string collectionPath, [ComAliasName("OLE.LPCOLESTR")]string propertyName, [ComAliasName("Microsoft.VisualStudio.Shell.Interop.VSENCLOSINGSCOPES")]out uint scopes) {
				scopes = (uint)inner.GetPropertyScopes(collectionPath, propertyName);
				return 0;
			}

			public int GetReadOnlySettingsStore([ComAliasName("Microsoft.VisualStudio.Shell.Interop.VSSETTINGSSCOPE")]uint scope, out IVsSettingsStore store) {
				store = new StoreWrapper(inner.GetReadOnlySettingsStore((SettingsScope)scope));
				return 0;
			}

			public int GetWritableSettingsStore([ComAliasName("Microsoft.VisualStudio.Shell.Interop.VSSETTINGSSCOPE")]uint scope, out IVsWritableSettingsStore writableStore) {
				writableStore = (IVsWritableSettingsStore)inner.GetReadOnlySettingsStore((SettingsScope)scope);
				return 0;
			}
		}

		class StoreWrapper : IVsSettingsStore {
			readonly SettingsStore inner;

			public StoreWrapper(SettingsStore inner) {
				this.inner = inner;
			}

			public int CollectionExists([ComAliasName("OLE.LPCOLESTR")]string collectionPath, [ComAliasName("OLE.BOOL")]out int pfExists) {
				pfExists = inner.CollectionExists(collectionPath) ? 1 : 0;
				return 0;
			}

			[SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "No native resources are involved.")]
			public int GetBinary([ComAliasName("OLE.LPCOLESTR")]string collectionPath, [ComAliasName("OLE.LPCOLESTR")]string propertyName, [ComAliasName("OLE.ULONG")]uint byteLength, [ComAliasName("TextManager.BYTE")]byte[] pBytes, [ComAliasName("OLE.ULONG")]uint[] actualByteLength) {
				var stream = inner.GetMemoryStream(collectionPath, propertyName);
				if (byteLength == 0 || pBytes == null)
					actualByteLength[0] = (uint)stream.Length;
				else
					stream.CopyTo(new MemoryStream(pBytes, 0, (int)byteLength));
				return 0;
			}

			public int GetBool([ComAliasName("OLE.LPCOLESTR")]string collectionPath, [ComAliasName("OLE.LPCOLESTR")]string propertyName, [ComAliasName("OLE.BOOL")]out int value) {
				value = inner.GetBoolean(collectionPath, propertyName) ? 1 : 0;
				return 0;
			}

			public int GetBoolOrDefault([ComAliasName("OLE.LPCOLESTR")]string collectionPath, [ComAliasName("OLE.LPCOLESTR")]string propertyName, [ComAliasName("OLE.BOOL")]int defaultValue, [ComAliasName("OLE.BOOL")]out int value) {
				value = inner.GetBoolean(collectionPath, propertyName, defaultValue != 0) ? 1 : 0;
				return 0;
			}

			public int GetInt([ComAliasName("OLE.LPCOLESTR")]string collectionPath, [ComAliasName("OLE.LPCOLESTR")]string propertyName, out int value) {
				value = inner.GetInt32(collectionPath, propertyName);
				return 0;
			}

			public int GetInt64([ComAliasName("OLE.LPCOLESTR")]string collectionPath, [ComAliasName("OLE.LPCOLESTR")]string propertyName, out long value) {
				value = inner.GetInt64(collectionPath, propertyName);
				return 0;
			}

			public int GetInt64OrDefault([ComAliasName("OLE.LPCOLESTR")]string collectionPath, [ComAliasName("OLE.LPCOLESTR")]string propertyName, long defaultValue, out long value) {
				value = inner.GetInt64(collectionPath, propertyName, defaultValue);
				return 0;
			}

			public int GetIntOrDefault([ComAliasName("OLE.LPCOLESTR")]string collectionPath, [ComAliasName("OLE.LPCOLESTR")]string propertyName, int defaultValue, out int value) {
				value = inner.GetInt32(collectionPath, propertyName, defaultValue);
				return 0;
			}

			public int GetLastWriteTime([ComAliasName("OLE.LPCOLESTR")]string collectionPath, [ComAliasName("VsShell.SYSTEMTIME")]SYSTEMTIME[] lastWriteTime) {
				var dt = inner.GetLastWriteTime(collectionPath);
				lastWriteTime[0].wDay = (ushort)dt.Day;
				lastWriteTime[0].wDayOfWeek = (ushort)dt.DayOfWeek;
				lastWriteTime[0].wHour = (ushort)dt.Hour;
				lastWriteTime[0].wMilliseconds = (ushort)dt.Millisecond;
				lastWriteTime[0].wMinute = (ushort)dt.Minute;
				lastWriteTime[0].wMonth = (ushort)dt.Month;
				lastWriteTime[0].wSecond = (ushort)dt.Second;
				lastWriteTime[0].wYear = (ushort)dt.Year;
				return 0;
			}

			public int GetPropertyCount([ComAliasName("OLE.LPCOLESTR")]string collectionPath, [ComAliasName("OLE.DWORD")]out uint propertyCount) {
				propertyCount = (uint)inner.GetPropertyCount(collectionPath);
				return 0;
			}

			public int GetPropertyName([ComAliasName("OLE.LPCOLESTR")]string collectionPath, [ComAliasName("OLE.DWORD")]uint index, out string propertyName) {
				propertyName = inner.GetPropertyNames(collectionPath).ElementAt((int)index);
				return 0;
			}

			public int GetPropertyType([ComAliasName("OLE.LPCOLESTR")]string collectionPath, [ComAliasName("OLE.LPCOLESTR")]string propertyName, [ComAliasName("Microsoft.VisualStudio.Shell.Interop.VSSETTINGSTYPE")]out uint type) {
				type = (uint)inner.GetPropertyType(collectionPath, propertyName);
				return 0;
			}

			public int GetString([ComAliasName("OLE.LPCOLESTR")]string collectionPath, [ComAliasName("OLE.LPCOLESTR")]string propertyName, out string value) {
				value = inner.GetString(collectionPath, propertyName);
				return 0;
			}

			public int GetStringOrDefault([ComAliasName("OLE.LPCOLESTR")]string collectionPath, [ComAliasName("OLE.LPCOLESTR")]string propertyName, [ComAliasName("OLE.LPCOLESTR")]string defaultValue, out string value) {
				if (defaultValue == null)
					value = inner.GetString(collectionPath, propertyName);
				else
					value = inner.GetString(collectionPath, propertyName, defaultValue);
				return 0;
			}

			public int GetSubCollectionCount([ComAliasName("OLE.LPCOLESTR")]string collectionPath, [ComAliasName("OLE.DWORD")]out uint subCollectionCount) {
				subCollectionCount = (uint)inner.GetSubCollectionCount(collectionPath);
				return 0;
			}

			public int GetSubCollectionName([ComAliasName("OLE.LPCOLESTR")]string collectionPath, [ComAliasName("OLE.DWORD")]uint index, out string subCollectionName) {
				subCollectionName = inner.GetSubCollectionNames(collectionPath).ElementAt((int)index);
				return 0;
			}

			public int GetUnsignedInt([ComAliasName("OLE.LPCOLESTR")]string collectionPath, [ComAliasName("OLE.LPCOLESTR")]string propertyName, [ComAliasName("OLE.DWORD")]out uint value) {
				value = inner.GetUInt32(collectionPath, propertyName);
				return 0;
			}

			public int GetUnsignedInt64([ComAliasName("OLE.LPCOLESTR")]string collectionPath, [ComAliasName("OLE.LPCOLESTR")]string propertyName, out ulong value) {
				value = inner.GetUInt64(collectionPath, propertyName);
				return 0;
			}

			public int GetUnsignedInt64OrDefault([ComAliasName("OLE.LPCOLESTR")]string collectionPath, [ComAliasName("OLE.LPCOLESTR")]string propertyName, ulong defaultValue, out ulong value) {
				value = inner.GetUInt64(collectionPath, propertyName, defaultValue);
				return 0;
			}

			public int GetUnsignedIntOrDefault([ComAliasName("OLE.LPCOLESTR")]string collectionPath, [ComAliasName("OLE.LPCOLESTR")]string propertyName, [ComAliasName("OLE.DWORD")]uint defaultValue, [ComAliasName("OLE.DWORD")]out uint value) {
				value = inner.GetUInt32(collectionPath, propertyName, defaultValue);
				return 0;
			}

			public int PropertyExists([ComAliasName("OLE.LPCOLESTR")]string collectionPath, [ComAliasName("OLE.LPCOLESTR")]string propertyName, [ComAliasName("OLE.BOOL")]out int pfExists) {
				pfExists = inner.PropertyExists(collectionPath, propertyName) ? 1 : 0;
				return 0;
			}
		}

		class DummyLog : IVsActivityLog {
			public int LogEntry([ComAliasName("Microsoft.VisualStudio.Shell.Interop.ACTIVITYLOG_ENTRYTYPE")]uint actType, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszSource, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszDescription) {
				return 0;
			}

			public int LogEntryGuid([ComAliasName("Microsoft.VisualStudio.Shell.Interop.ACTIVITYLOG_ENTRYTYPE")]uint actType, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszSource, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszDescription, Guid guid) {
				return 0;
			}

			public int LogEntryGuidHr([ComAliasName("Microsoft.VisualStudio.Shell.Interop.ACTIVITYLOG_ENTRYTYPE")]uint actType, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszSource, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszDescription, Guid guid, int hr) {
				return 0;
			}

			public int LogEntryGuidHrPath([ComAliasName("Microsoft.VisualStudio.Shell.Interop.ACTIVITYLOG_ENTRYTYPE")]uint actType, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszSource, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszDescription, Guid guid, int hr, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszPath) {
				return 0;
			}

			public int LogEntryGuidPath([ComAliasName("Microsoft.VisualStudio.Shell.Interop.ACTIVITYLOG_ENTRYTYPE")]uint actType, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszSource, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszDescription, Guid guid, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszPath) {
				return 0;
			}

			public int LogEntryHr([ComAliasName("Microsoft.VisualStudio.Shell.Interop.ACTIVITYLOG_ENTRYTYPE")]uint actType, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszSource, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszDescription, int hr) {
				return 0;
			}

			public int LogEntryHrPath([ComAliasName("Microsoft.VisualStudio.Shell.Interop.ACTIVITYLOG_ENTRYTYPE")]uint actType, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszSource, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszDescription, int hr, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszPath) {
				return 0;
			}

			public int LogEntryPath([ComAliasName("Microsoft.VisualStudio.Shell.Interop.ACTIVITYLOG_ENTRYTYPE")]uint actType, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszSource, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszDescription, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszPath) {
				return 0;
			}
		}

		class DummyWindowManager : IVsWindowManager2 {
			[return: MarshalAs(UnmanagedType.IUnknown)]
			public object GetResourceKeyReferenceType([In, MarshalAs(UnmanagedType.IUnknown)]object requestedResource) {
				Type left = requestedResource as Type;
				if (left == null) {
					throw new ArgumentException();
				}
				if (left == typeof(ScrollBar) || left == typeof(ScrollViewer)) {
					return typeof(MainWindow);
				}
				throw new ArgumentException();
			}

			public void _VtblGap1_3() { }
		}
	}

	class DummyVsMonitorSelection : IVsMonitorSelection {
		public int AdviseSelectionEvents(IVsSelectionEvents pSink, out uint pdwCookie) {
			pdwCookie = 0;
			return 0;
		}

		public int GetCmdUIContextCookie(ref Guid rguidCmdUI, out uint pdwCmdUICookie) {
			pdwCmdUICookie = 0;
			return 0;
		}

		public int GetCurrentElementValue(uint elementid, out object pvarValue) {
			pvarValue = null;
			return 0;
		}

		public int GetCurrentSelection(out IntPtr ppHier, out uint pitemid, out IVsMultiItemSelect ppMIS, out IntPtr ppSC) {
			pitemid = 0;
			ppHier = IntPtr.Zero;
			ppMIS = null;
			ppSC = IntPtr.Zero;
			return 0;
		}

		public int IsCmdUIContextActive(uint dwCmdUICookie, out int pfActive) {
			pfActive = 0;
			return 0;
		}

		public int SetCmdUIContext(uint dwCmdUICookie, int fActive) {
			return 0;
		}

		public int UnadviseSelectionEvents(uint dwCookie) {
			return 0;
		}
	}
}

namespace Microsoft.Internal.VisualStudio.Shell.Interop {
	[CompilerGenerated, Guid("D73DC67C-3E91-4073-9A5E-5D09AA74529B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), TypeIdentifier]
	[ComImport]
	public interface IVsWindowManager2 {
		void _VtblGap1_3();
		[MethodImpl(MethodImplOptions.InternalCall)]
		[return: MarshalAs(UnmanagedType.IUnknown)]
		object GetResourceKeyReferenceType([MarshalAs(UnmanagedType.IUnknown)] [In] object requestedResource);
	}
}
