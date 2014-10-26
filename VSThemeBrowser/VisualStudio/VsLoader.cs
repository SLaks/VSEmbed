using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace VSThemeBrowser.VisualStudio {
	static class VsLoader {
		public static IEnumerable<Version> FindVsVersions() {
			// Grab every version of every SKU, sorted by SKU priority
			// Then, filter out duplicate versions in later SKUs only.
			return SkuKeyNames
				.SelectMany(GetSkuVersions)
				.Distinct()
				.OrderBy(d => d);
		}
		private static IEnumerable<Version> GetSkuVersions(string sku) {
			using (var software = Registry.LocalMachine.OpenSubKey("SOFTWARE"))
			using (var ms = software.OpenSubKey("Microsoft"))
			using (var vs = ms.OpenSubKey(sku))
				return vs == null ? new Version[0] : vs.GetSubKeyNames()
						.Select(s => {
							Version v;
							Version.TryParse(s, out v);
							return v;
						})
						.Where(d => d != null);
		}


		/// <summary>
		/// A list of key names for versions of Visual Studio which have the editor components 
		/// necessary to create an EditorHost instance.  Listed in preference order.
		/// Stolen from @JaredPar
		/// </summary>
		private static readonly string[] SkuKeyNames = {
			"VisualStudio",	// Standard non-express SKU of Visual Studio
			"WDExpress",	// Windows Desktop express
			"VCSExpress",	// Visual C# express
			"VCExpress",	// Visual C++ express
			"VBExpress",	// Visual Basic Express
		};

		public static string GetVersionPath(Version version) {
			return SkuKeyNames.Select(sku =>
				Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\" + sku + @"\" + version.ToString(2), "InstallDir", null) as string
			).FirstOrDefault(p => p != null);
		}

		public static void InitializeLatest() {
			Initialize(FindVsVersions().Last());
		}

		public static void Initialize(Version vsVersion) {
			if (VsVersion != null)
				throw new InvalidOperationException("VsLoader cannot be initialized twice");
			VsVersion = vsVersion;
			TryLoadInteropAssembly(GetVersionPath(vsVersion));
			AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
        }

		///<summary>The version of Visual Studio that will be loaded.  This cannot be changed, because the CLR caches assembly loads.</summary>
		public static Version VsVersion { get; private set; }

		static readonly Regex versionMatcher = new Regex(@"(?<=\.)\d+\.0$");
		static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args) {
			if (!args.Name.StartsWith("Microsoft.VisualStudio"))
				return null;

			var name = new AssemblyName(args.Name);
			if (name.Version != null && name.Version.Major == VsVersion.Major)
				return null;	// Don't recurse.  I check the major version only because AssemblyName will resolve the build number from the GAC.

			// Always specify a complete version to avoid partial assembly loading, which skips the GAC.
			name.Version = new Version(VsVersion.Major, VsVersion.Minor, 0, 0);
			name.Name = versionMatcher.Replace(name.Name, VsVersion.ToString(2));
			Debug.WriteLine("Redirecting load of " + args.Name + ",\tfrom " + (args.RequestingAssembly == null ? "(unknown)" : args.RequestingAssembly.FullName));

			try {
				return Assembly.Load(name);
			} catch (FileNotFoundException) {
				var directory = GetVersionPath(VsVersion);
				if (name.Name.EndsWith(".resources"))
					return LoadResourceDll(name, directory, name.CultureInfo)
						?? LoadResourceDll(name, directory, name.CultureInfo.Parent);
				return Assembly.LoadFile(Path.Combine(directory, name.Name + ".dll"));
			}
		}
		static Assembly LoadResourceDll(AssemblyName name, string baseDirectory, CultureInfo culture) {
			var dllPath = Path.Combine(baseDirectory, culture.Name, name.Name + ".dll");
			if (!File.Exists(dllPath))
				return null;
			return Assembly.LoadFile(dllPath);
		}

		/// <summary>
		/// The interop assembly isn't included in the GAC and it doesn't offer any MEF components (it's
		/// just a simple COM interop library).  Hence it needs to be loaded a bit specially.  Just find
		/// the assembly on disk and hook into the resolve event.
		/// Copied from @JaredPar's EditorUtils.
		/// </summary>
		private static bool TryLoadInteropAssembly(string installDirectory) {
			const string interopName = "Microsoft.VisualStudio.Platform.VSEditor.Interop";
			const string interopNameWithExtension = interopName + ".dll";
			var interopAssemblyPath = Path.Combine(installDirectory, "PrivateAssemblies");
			interopAssemblyPath = Path.Combine(interopAssemblyPath, interopNameWithExtension);
			try {
				var interopAssembly = Assembly.LoadFile(interopAssemblyPath);
				if (interopAssembly == null) {
					return false;
				}

				var comparer = StringComparer.OrdinalIgnoreCase;
				AppDomain.CurrentDomain.AssemblyResolve += (sender, e) => {
					if (comparer.Equals(e.Name, interopAssembly.FullName)) {
						return interopAssembly;
					}
					return null;
				};

				return true;
			} catch (Exception) {
				return false;
			}
		}
	}
}
