using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell.Interop;

namespace VSEmbed.Services {
	///<summary>Wraps the (half-)managed ExternalSettingsManager implementation in an IVsSettingsManager so that it can be consumed by code that expects the native interface.</summary>
	class SettingsManagerWrapper : IVsSettingsManager, IDisposable {
		readonly ExternalSettingsManager inner;

		public SettingsManagerWrapper(ExternalSettingsManager inner) {
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
	}
}