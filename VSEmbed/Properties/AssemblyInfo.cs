using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("VSEmbed")]
[assembly: AssemblyDescription("Embeds Visual Studio's editor and theming system in standalone projects.")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("SLaks")]
[assembly: AssemblyProduct("VSEmbed")]
[assembly: AssemblyCopyright("Copyright © SLaks 2014")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("93bbeb44-6849-423e-a5c1-440558661b5c")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Build and Revision Numbers 
// by using the '*' as shown below:
// [assembly: AssemblyVersion("1.0.*")]

[assembly: AssemblyVersion(ProductVersion.Current)]
[assembly: AssemblyFileVersion(ProductVersion.Current)]
[assembly: AssemblyInformationalVersion(ProductVersion.Current)]
static class ProductVersion { public const string Current = "1.2.0"; }

