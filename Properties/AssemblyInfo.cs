// This is the correct code for your Properties/AssemblyInfo.cs file.
// The conflicting AssemblyVersion line has been removed.

using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Rhino.PlugIns;

// Plug-in Description Attributes - all of these are optional.
// These will show in Rhino's option dialog, in the tab Plug-ins.
//
// NOTE: The AssemblyVersion attribute has been removed from this file.
// It is now managed automatically by the project properties (the .csproj file).
// You can set the version by right-clicking the project in Solution Explorer -> Properties -> Package -> Package version.

[assembly: PlugInDescription(DescriptionType.Address, "Lindenspürstraße 32, 70176 Stuttgart")]
[assembly: PlugInDescription(DescriptionType.Country, "Germany")]
[assembly: PlugInDescription(DescriptionType.Email, "shrestha@str-ucture.com")]
[assembly: PlugInDescription(DescriptionType.Phone, "+49-1742738566")]
[assembly: PlugInDescription(DescriptionType.Fax, "")]
[assembly: PlugInDescription(DescriptionType.Organization, "str.ucture GmbH")]
[assembly: PlugInDescription(DescriptionType.UpdateUrl, "")]
[assembly: PlugInDescription(DescriptionType.WebSite, "")]

// Icons should be Windows .ico files and contain 32-bit images in the following sizes: 16, 24, 32, 48, and 256.
[assembly: PlugInDescription(DescriptionType.Icon, "MyProject.Resources.srtLCA32.ico")]

// The following GUID is for the ID of the typelib if this project is exposed to COM
// This will also be the Guid of the Rhino plug-in
[assembly: Guid("01D22D46-3C89-43A5-A92B-A1D6519A4244")]