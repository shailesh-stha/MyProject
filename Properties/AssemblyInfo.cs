using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Rhino.PlugIns;

// Plug-in Description Attributes - all of these are optional.
// These will show in Rhino's option dialog, in the tab Plug-ins.
[assembly: PlugInDescription(DescriptionType.Address, "Lindenspürstraße 32, 70176 Stuttgart")]
[assembly: PlugInDescription(DescriptionType.Country, "Germany")]
[assembly: PlugInDescription(DescriptionType.Email, "shrestha@str-ucture.com")]
[assembly: PlugInDescription(DescriptionType.Phone, "+49-1742738566")]
[assembly: PlugInDescription(DescriptionType.Fax, "")]
[assembly: PlugInDescription(DescriptionType.Organization, "str.ucture GmbH")]
[assembly: PlugInDescription(DescriptionType.UpdateUrl, "")]
[assembly: PlugInDescription(DescriptionType.WebSite, "")]

// Icons should be Windows .ico files and contain 32-bit images in the following sizes: 16, 24, 32, 48, and 256.
[assembly: PlugInDescription(DescriptionType.Icon, "MyProject.EmbeddedResources.strLCA32.ico")]

// The following GUID is for the ID of the typelib if this project is exposed to COM
// This will also be the Guid of the Rhino plug-in
[assembly: Guid("4049cb58-d5dd-4d34-8cf8-8e06137ead6e")] 
