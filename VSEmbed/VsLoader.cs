using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;

namespace VSEmbed {
	///<summary>Sets up assembly redirection to load Visual Studio assemblies.</summary>
	///<remarks>This class must be initialized before anything else is JITted.</remarks>
	public static class VsLoader {
		///<summary>Finds all installed Visual Studio versions.</summary>
		public static IEnumerable<Version> FindAllVersions() {
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

		///<summary>Gets the installation directory for the specified version.</summary>
		private static string GetInstallationDirectory(Version version) {
			return SkuKeyNames.Select(sku =>
				Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\" + sku + @"\" + version.ToString(2), "InstallDir", null) as string
			).FirstOrDefault(p => p != null);
		}

		///<summary>Indicates whether the code is running within the VS designer.</summary>
		public static bool IsDesignMode {
			get { return (bool)DesignerProperties.IsInDesignModeProperty.GetMetadata(typeof(DependencyObject)).DefaultValue; }
		}


		///<summary>Initializes the assembly loader with the latest installed version of Visual Studio.</summary>
		public static void LoadLatest() {
			Load(FindAllVersions().Last());
		}

		///<summary>Initializes the assembly loader with the specified version of Visual Studio.</summary>
		public static void Load(Version vsVersion) {
			if (VsVersion != null)
				throw new InvalidOperationException("VsLoader cannot be initialized twice");
			if (string.IsNullOrEmpty(GetInstallationDirectory(vsVersion)) || !Directory.Exists(GetInstallationDirectory(vsVersion)))
				throw new ArgumentException("Cannot locate Visual Studio v" + vsVersion);

			VsVersion = vsVersion;
			InstallationDirectory = GetInstallationDirectory(VsVersion);
			TryLoadInteropAssembly(InstallationDirectory);
			AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve_VS;

			if (RoslynAssemblyPath != null)
				AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve_Roslyn;
		}

		///<summary>Gets the version of Visual Studio that will be loaded.  This cannot be changed, because the CLR caches assembly loads.</summary>
		public static Version VsVersion { get; private set; }

		///<summary>Gets the installation directory for the loaded VS version.</summary>
		public static string InstallationDirectory { get; private set; }

		static readonly Regex versionMatcher = new Regex(@"(?<=\.)\d+\.0$");
		static Assembly CurrentDomain_AssemblyResolve_VS(object sender, ResolveEventArgs args) {
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
				if (name.Name.EndsWith(".resources"))
					return LoadResourceDll(name, InstallationDirectory, name.CultureInfo)
						?? LoadResourceDll(name, InstallationDirectory, name.CultureInfo.Parent);
				return Assembly.LoadFile(Path.Combine(InstallationDirectory, name.Name + ".dll"));
			}
		}
		static Assembly LoadResourceDll(AssemblyName name, string baseDirectory, CultureInfo culture) {
			var dllPath = Path.Combine(baseDirectory, culture.Name, name.Name + ".dll");
			if (!File.Exists(dllPath))
				return null;
			return Assembly.LoadFile(dllPath);
		}

		///<summary>Gets the directory containing Roslyn assemblies, or null if this VS version does not contain Roslyn.</summary>
		public static string RoslynAssemblyPath {
			get {
				// TODO: Use Roslyn Preview in Dev12?
				if (VsVersion.Major == 14)
					return Path.Combine(InstallationDirectory, "PrivateAssemblies");
				return null;	// TODO: Predict GAC / versioning for Dev15
			}
		}
		static readonly string[] RoslynAssemblyPrefixes = {
			"Microsoft.CodeAnalysis",
			"Roslyn.",  // For package assemblies like Roslyn.VisualStudio.Setup
			"System.Reflection.Metadata",
			"Microsoft.VisualStudio.LanguageServices",
			"Esent.Interop",
			"System.Composition.AttributedModel",		// New to VS2015 Preview
			"Microsoft.VisualStudio.Composition"		// For VS MEF in VS2015 Preview
		};
		static Assembly CurrentDomain_AssemblyResolve_Roslyn(object sender, ResolveEventArgs args) {
			if (!RoslynAssemblyPrefixes.Any(args.Name.StartsWith))
				return null;

			Debug.WriteLine("Redirecting load of " + args.Name + ",\tfrom " + (args.RequestingAssembly == null ? "(unknown)" : args.RequestingAssembly.FullName));

			var name = new AssemblyName(args.Name);
			var dllPath = Path.Combine(RoslynAssemblyPath, name.Name + ".dll");
			if (File.Exists(dllPath))
				return Assembly.LoadFile(dllPath);
			else if (name.Name.EndsWith(".resources"))
				return LoadResourceDll(name, RoslynAssemblyPath, name.CultureInfo)
					?? LoadResourceDll(name, RoslynAssemblyPath, name.CultureInfo.Parent);
			else
				return null;
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
			} catch {
				return false;
			}
		}
	}
}
