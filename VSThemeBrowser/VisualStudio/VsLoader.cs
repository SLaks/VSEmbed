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
			using (var software = Registry.LocalMachine.OpenSubKey("SOFTWARE"))
			using (var ms = software.OpenSubKey("Microsoft"))
			using (var vs = ms.OpenSubKey("VisualStudio"))
				return vs.GetSubKeyNames()
						.Select(s => {
							Version v;
							Version.TryParse(s, out v);
							return v;
						})
				.Where(d => d != null)
				.OrderBy(d => d);
		}

		public static string GetVersionExe(Version version) {
			return Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\VisualStudio\" + version.ToString(2) + @"\Setup\VS", "EnvironmentPath", null) as string;
		}

		public static void InitializeLatest() {
			Initialize(FindVsVersions().Last());
		}

		public static void Initialize(Version vsVersion) {
			if (VsVersion != null)
				throw new InvalidOperationException("VsLoader cannot be initialized twice");
			VsVersion = vsVersion;
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
				var directory = Path.GetDirectoryName(GetVersionExe(VsVersion));
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
	}
}
