// CsvReader.cs
using Rhino;
using System;
using System.Collections.Generic;
using System.Globalization;
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
        /// Reads custom attribute data, returning two independent lists of unique keys and values.
        /// </summary>
        public static void ReadCustomAttributeListsFromResource(string resourceName, out List<string> keys, out List<string> values)
        {
            var keySet = new HashSet<string>();
            var valueSet = new HashSet<string>();
            var assembly = Assembly.GetExecutingAssembly();

            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    RhinoApp.WriteLine($"Error: Embedded resource '{resourceName}' not found.");
                    keys = new List<string>();
                    values = new List<string>();
                    return;
                }

                using (var reader = new StreamReader(stream))
                {
                    reader.ReadLine(); // Skip header
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        var parts = line.Split(',');
                        if (parts.Length >= 1 && !string.IsNullOrWhiteSpace(parts[0]))
                        {
                            keySet.Add(parts[0].Trim());
                        }
                        if (parts.Length >= 2 && !string.IsNullOrWhiteSpace(parts[1]))
                        {
                            valueSet.Add(parts[1].Trim());
                        }
                    }
                }
            }
            keys = keySet.ToList();
            values = valueSet.ToList();
        }

        // ... (other methods remain unchanged) ...

        /// <summary>
        /// Reads material LCA data from an embedded CSV resource file. Returns a list to preserve order.
        /// </summary>
        public static List<MaterialData> ReadMaterialLcaDataFromResource(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    RhinoApp.WriteLine($"Error: Embedded resource '{resourceName}' not found.");
                    return new List<MaterialData>();
                }

                using (var reader = new StreamReader(stream))
                {
                    return ParseMaterialCsvContent(reader.ReadToEnd());
                }
            }
        }

        /// <summary>
        /// Dynamically reads material LCA data, preserving row order.
        /// </summary>
        public static List<MaterialData> ReadMaterialLcaDataDynamic(RhinoDoc doc)
        {
            if (doc != null)
            {
                var url = doc.Strings.GetValue("STR_MATERIAL_CSV_URL");
                if (!string.IsNullOrWhiteSpace(url))
                {
                    RhinoApp.WriteLine($"Attempting to load Material LCA data from custom URL: {url}");
                    try
                    {
                        string csvContent = Task.Run(async () => await GetContentFromUrlOrFile(url)).GetAwaiter().GetResult();

                        if (!string.IsNullOrEmpty(csvContent))
                        {
                            var list = ParseMaterialCsvContent(csvContent);
                            if (list.Any())
                            {
                                RhinoApp.WriteLine($"Successfully loaded {list.Count} materials from custom URL.");
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

            RhinoApp.WriteLine("Loading Material LCA data from default embedded resource.");
            return ReadMaterialLcaDataFromResource("MyProject.Resources.Data.materialListWithUnits.csv");
        }

        /// <summary>
        /// Parses a string of CSV content into a list of MaterialData to preserve order.
        /// </summary>
        private static List<MaterialData> ParseMaterialCsvContent(string csvContent)
        {
            var materials = new List<MaterialData>();
            using (var reader = new StringReader(csvContent))
            {
                reader.ReadLine(); // Skip header
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    var parts = line.Split(',');
                    if (parts.Length != 4) continue;

                    var materialName = parts[0].Trim();
                    if (double.TryParse(parts[1], NumberStyles.Any, CultureInfo.InvariantCulture, out double quantity) &&
                        double.TryParse(parts[3], NumberStyles.Any, CultureInfo.InvariantCulture, out double lcaValue) &&
                        !string.IsNullOrWhiteSpace(materialName))
                    {
                        materials.Add(new MaterialData
                        {
                            MaterialName = materialName,
                            ReferenceQuantity = quantity,
                            ReferenceUnit = parts[2].Trim(),
                            Lca = lcaValue
                        });
                    }
                }
            }
            return materials;
        }

        /// <summary>
        /// Dynamically reads hierarchical IFC classes, preserving row order.
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

            RhinoApp.WriteLine("Loading IFC classes from default embedded resource.");
            return ReadIfcClassesFromResource("MyProject.Resources.Data.ifcClassListWithSubClass.csv");
        }

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
            return ifcClasses;
        }

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