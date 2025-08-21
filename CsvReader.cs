// CsvReader.cs
using Rhino;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Threading.Tasks;

namespace MyProject
{
    /// <summary>
    /// A utility class to read data from various sources.
    /// </summary>
    public static class CsvReader
    {
        /// <summary>
        /// Reads material LCA data from an embedded CSV resource file.
        /// </summary>
        public static Dictionary<string, MaterialData> ReadMaterialLcaDataFromResource(string resourceName)
        {
            var materials = new Dictionary<string, MaterialData>();
            var assembly = Assembly.GetExecutingAssembly();

            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    RhinoApp.WriteLine($"Error: Embedded resource '{resourceName}' not found.");
                    return materials;
                }

                using (var reader = new StreamReader(stream))
                {
                    reader.ReadLine(); // Skip header
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        var parts = line.Split(',');
                        if (parts.Length != 4) continue;

                        var materialName = parts[0].Trim();
                        if (double.TryParse(parts[1], out double quantity) &&
                            double.TryParse(parts[3], out double lcaValue) &&
                            !string.IsNullOrWhiteSpace(materialName))
                        {
                            materials[materialName] = new MaterialData
                            {
                                MaterialName = materialName,
                                ReferenceQuantity = quantity,
                                ReferenceUnit = parts[2].Trim(),
                                Lca = lcaValue
                            };
                        }
                    }
                }
            }
            return materials;
        }

        /// <summary>
        /// Dynamically reads hierarchical IFC classes based on the active Rhino document's settings.
        /// The CSV is expected to have two columns: PrimaryClass,SecondaryClass.
        /// </summary>
        public static Dictionary<string, List<string>> ReadIfcClassesDynamic(RhinoDoc doc)
        {
            if (doc != null)
            {
                var url = doc.Strings.GetValue("STR_IFC_CLASS_CSV_URL");
                if (!string.IsNullOrWhiteSpace(url))
                {
                    RhinoApp.WriteLine($"Attempting to load IFC classes from custom URL: {url}");
                    try
                    {
                        string csvContent = Task.Run(async () => await GetContentFromUrlOrFile(url)).GetAwaiter().GetResult();

                        if (!string.IsNullOrEmpty(csvContent))
                        {
                            var dict = ParseIfcCsvContent(csvContent);
                            if (dict.Any())
                            {
                                RhinoApp.WriteLine($"Successfully loaded {dict.Count} primary IFC classes from custom URL.");
                                return dict;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        RhinoApp.WriteLine($"Error loading from custom URL '{url}'. Details: {ex.Message}");
                    }
                }
            }

            // Fallback to embedded resource
            RhinoApp.WriteLine("Loading IFC classes from default embedded resource.");
            // Assuming the fallback resource is also in the new format.
            // You might need to update the resource file in your project.
            return ReadIfcClassesFromResource("MyProject.Resources.Data.ifcClassListWithSubClass.csv");
        }

        /// <summary>
        /// Helper to get CSV content from a web URL or a local file path.
        /// </summary>
        private static async Task<string> GetContentFromUrlOrFile(string url)
        {
            if (Uri.TryCreate(url, UriKind.Absolute, out Uri webUri) && (webUri.Scheme == Uri.UriSchemeHttp || webUri.Scheme == Uri.UriSchemeHttps))
            {
                try
                {
                    using (var client = new HttpClient())
                    {
                        return await client.GetStringAsync(webUri);
                    }
                }
                catch (HttpRequestException)
                {
                    // Fall through to try as a local file
                }
            }

            if (File.Exists(url))
            {
                return File.ReadAllText(url);
            }

            throw new FileNotFoundException($"The specified resource could not be found as a web URL or local file path.", url);
        }

        /// <summary>
        /// Parses a string of CSV content into a dictionary of IFC classes and subclasses.
        /// </summary>
        private static Dictionary<string, List<string>> ParseIfcCsvContent(string csvContent)
        {
            var ifcClasses = new Dictionary<string, List<string>>();
            using (var reader = new StringReader(csvContent))
            {
                reader.ReadLine(); // Skip header
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    var parts = line.Split(',');
                    if (parts.Length < 1 || string.IsNullOrWhiteSpace(parts[0])) continue;

                    string primaryClass = parts[0].Trim();
                    string secondaryClass = parts.Length > 1 ? parts[1].Trim() : string.Empty;

                    if (!ifcClasses.ContainsKey(primaryClass))
                    {
                        ifcClasses[primaryClass] = new List<string>();
                    }

                    if (!string.IsNullOrEmpty(secondaryClass) && !ifcClasses[primaryClass].Contains(secondaryClass))
                    {
                        ifcClasses[primaryClass].Add(secondaryClass);
                    }
                }
            }
            // Ensure subclasses are sorted for consistent UI
            foreach (var key in ifcClasses.Keys)
            {
                ifcClasses[key].Sort();
            }
            return ifcClasses;
        }

        /// <summary>
        /// Reads IFC classes and subclasses from an embedded CSV resource file. (Fallback method)
        /// </summary>
        public static Dictionary<string, List<string>> ReadIfcClassesFromResource(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    RhinoApp.WriteLine($"Error: Embedded resource '{resourceName}' not found.");
                    return new Dictionary<string, List<string>>();
                }
                using (var reader = new StreamReader(stream))
                {
                    return ParseIfcCsvContent(reader.ReadToEnd());
                }
            }
        }
    }
}