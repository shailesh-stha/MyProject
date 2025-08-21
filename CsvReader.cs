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
        /// Dynamically reads IFC classes based on the active Rhino document's settings.
        /// Order of precedence:
        /// 1. URL from Document User Text key "STR_IFC_CLASS_CSV_URL" (Web or Local).
        /// 2. Embedded resource file as a fallback.
        /// </summary>
        public static List<string> ReadIfcClassesDynamic(RhinoDoc doc)
        {
            if (doc != null)
            {
                var url = doc.Strings.GetValue("STR_IFC_CLASS_CSV_URL");
                if (!string.IsNullOrWhiteSpace(url))
                {
                    RhinoApp.WriteLine($"Attempting to load IFC classes from custom URL: {url}");
                    try
                    {
                        // Await the async method and get the result.
                        string csvContent = Task.Run(async () => await GetContentFromUrlOrFile(url)).GetAwaiter().GetResult();

                        if (!string.IsNullOrEmpty(csvContent))
                        {
                            var list = ParseIfcCsvContent(csvContent);
                            if (list.Any())
                            {
                                RhinoApp.WriteLine($"Successfully loaded {list.Count} IFC classes from custom URL.");
                                return list;
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
            return ReadIfcClassesFromResource("MyProject.Resources.Data.ifcClassList.csv");
        }

        /// <summary>
        /// Helper to get CSV content from a web URL or a local file path.
        /// </summary>
        private static async Task<string> GetContentFromUrlOrFile(string url)
        {
            // Try as a web URL first
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

            // If it's not a web URL or if web failed, try as a local file
            if (File.Exists(url))
            {
                return File.ReadAllText(url);
            }

            throw new FileNotFoundException($"The specified resource could not be found as a web URL or local file path.", url);
        }

        /// <summary>
        /// Parses a string of CSV content into a list of IFC classes.
        /// </summary>
        private static List<string> ParseIfcCsvContent(string csvContent)
        {
            var ifcClasses = new List<string>();
            using (var reader = new StringReader(csvContent))
            {
                reader.ReadLine(); // Skip header
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    var parts = line.Split(',');
                    if (parts.Length > 0 && !string.IsNullOrWhiteSpace(parts[0]))
                    {
                        ifcClasses.Add(parts[0].Trim());
                    }
                }
            }
            return ifcClasses;
        }

        /// <summary>
        /// Reads IFC classes from an embedded CSV resource file. (Fallback method)
        /// </summary>
        public static List<string> ReadIfcClassesFromResource(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    RhinoApp.WriteLine($"Error: Embedded resource '{resourceName}' not found.");
                    return new List<string>();
                }
                using (var reader = new StreamReader(stream))
                {
                    return ParseIfcCsvContent(reader.ReadToEnd());
                }
            }
        }
    }
}