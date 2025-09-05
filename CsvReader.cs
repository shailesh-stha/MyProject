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
        // Private helper methods for parsing and loading content
        private static (List<string> keys, List<string> values) ParseCustomCsvContent(string csvContent)
        {
            var keySet = new HashSet<string>();
            var valueSet = new HashSet<string>();
            using (var reader = new StringReader(csvContent))
            {
                reader.ReadLine(); // Skip header
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    var parts = line.Split(',');
                    if (parts.Length >= 1 && !string.IsNullOrWhiteSpace(parts[0])) keySet.Add(parts[0].Trim());
                    if (parts.Length >= 2 && !string.IsNullOrWhiteSpace(parts[1])) valueSet.Add(parts[1].Trim());
                }
            }
            return (keySet.ToList(), valueSet.ToList());
        }

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
                    if (!ifcClasses.ContainsKey(primaryClass)) ifcClasses[primaryClass] = new List<string>();
                    if (!string.IsNullOrEmpty(secondaryClass) && !ifcClasses[primaryClass].Contains(secondaryClass))
                    {
                        ifcClasses[primaryClass].Add(secondaryClass);
                    }
                }
            }
            return ifcClasses;
        }

        private static async Task<string> GetContentFromUrlOrFile(string url)
        {
            if (!string.IsNullOrWhiteSpace(url)) url = url.Trim('"');
            if (Uri.TryCreate(url, UriKind.Absolute, out Uri webUri) && (webUri.Scheme == Uri.UriSchemeHttp || webUri.Scheme == Uri.UriSchemeHttps))
            {
                try
                {
                    using (var client = new HttpClient())
                    {
                        return await client.GetStringAsync(webUri);
                    }
                }
                catch (HttpRequestException) { /* Fall through to try as a local file */ }
            }
            if (File.Exists(url)) return File.ReadAllText(url);
            throw new FileNotFoundException($"The specified resource could not be found as a web URL or local file path.", url);
        }

        // Tiered loading logic for each data type
        private static (List<string> keys, List<string> values)? TryLoadCustomDataFromPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return null;
            try
            {
                string csvContent = Task.Run(async () => await GetContentFromUrlOrFile(path)).GetAwaiter().GetResult();
                if (!string.IsNullOrEmpty(csvContent))
                {
                    var result = ParseCustomCsvContent(csvContent);
                    if (result.keys.Any() || result.values.Any())
                    {
                        RhinoApp.WriteLine($"Successfully loaded {result.keys.Count} custom keys and {result.values.Count} custom values from: {path}");
                        return result;
                    }
                }
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Failed to load from path '{path}'. Error: {ex.Message}");
            }
            return null;
        }

        private static List<MaterialData> TryLoadMaterialDataFromPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return null;
            try
            {
                string csvContent = Task.Run(async () => await GetContentFromUrlOrFile(path)).GetAwaiter().GetResult();
                if (!string.IsNullOrEmpty(csvContent))
                {
                    var list = ParseMaterialCsvContent(csvContent);
                    if (list.Any())
                    {
                        RhinoApp.WriteLine($"Successfully loaded {list.Count} materials from: {path}");
                        return list;
                    }
                }
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Failed to load from path '{path}'. Error: {ex.Message}");
            }
            return null;
        }

        private static Dictionary<string, List<string>> TryLoadIfcDataFromPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return null;
            try
            {
                string csvContent = Task.Run(async () => await GetContentFromUrlOrFile(path)).GetAwaiter().GetResult();
                if (!string.IsNullOrEmpty(csvContent))
                {
                    var dict = ParseIfcCsvContent(csvContent);
                    if (dict.Any())
                    {
                        RhinoApp.WriteLine($"Successfully loaded {dict.Count} primary IFC classes from: {path}");
                        return dict;
                    }
                }
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Failed to load from path '{path}'. Error: {ex.Message}");
            }
            return null;
        }

        // Public methods for dynamic loading
        public static void ReadCustomAttributeListsDynamic(RhinoDoc doc, out List<string> keys, out List<string> values)
        {
            var docPath = doc?.Strings.GetValue("STR_CUSTOM_CSV_URL");
            if (!string.IsNullOrWhiteSpace(docPath))
            {
                RhinoApp.WriteLine($"Attempting to load Custom Definitions from document path: {docPath}");
                var result = TryLoadCustomDataFromPath(docPath);
                if (result.HasValue)
                {
                    keys = result.Value.keys;
                    values = result.Value.values;
                    return;
                }
            }

            var globalPath = PluginSettingsManager.GetString(SettingKeys.CustomCsvPath, string.Empty);
            if (!string.IsNullOrWhiteSpace(globalPath))
            {
                RhinoApp.WriteLine($"Attempting to load Custom Definitions from global settings path: {globalPath}");
                var result = TryLoadCustomDataFromPath(globalPath);
                if (result.HasValue)
                {
                    keys = result.Value.keys;
                    values = result.Value.values;
                    return;
                }
            }

            RhinoApp.WriteLine("Loading Custom Definitions from default embedded resource.");
            ReadCustomAttributeListsFromResource("MyProject.Resources.Data.customAttribute.csv", out keys, out values);
        }

        public static List<MaterialData> ReadMaterialLcaDataDynamic(RhinoDoc doc)
        {
            var docPath = doc?.Strings.GetValue("STR_MATERIAL_CSV_URL");
            if (!string.IsNullOrWhiteSpace(docPath))
            {
                RhinoApp.WriteLine($"Attempting to load Material LCA data from document path: {docPath}");
                var materialList = TryLoadMaterialDataFromPath(docPath);
                if (materialList != null) return materialList;
            }

            var globalPath = PluginSettingsManager.GetString(SettingKeys.MaterialCsvPath, string.Empty);
            if (!string.IsNullOrWhiteSpace(globalPath))
            {
                RhinoApp.WriteLine($"Attempting to load Material LCA data from global settings path: {globalPath}");
                var materialList = TryLoadMaterialDataFromPath(globalPath);
                if (materialList != null) return materialList;
            }

            RhinoApp.WriteLine("Loading Material LCA data from default embedded resource.");
            return ReadMaterialLcaDataFromResource("MyProject.Resources.Data.materialListWithUnits.csv");
        }

        public static Dictionary<string, List<string>> ReadIfcClassesDynamic(RhinoDoc doc)
        {
            var docPath = doc?.Strings.GetValue("STR_IFC_CLASS_CSV_URL");
            if (!string.IsNullOrWhiteSpace(docPath))
            {
                RhinoApp.WriteLine($"Attempting to load IFC classes from document path: {docPath}");
                var ifcDict = TryLoadIfcDataFromPath(docPath);
                if (ifcDict != null) return ifcDict;
            }

            var globalPath = PluginSettingsManager.GetString(SettingKeys.IfcClassCsvPath, string.Empty);
            if (!string.IsNullOrWhiteSpace(globalPath))
            {
                RhinoApp.WriteLine($"Attempting to load IFC classes from global settings path: {globalPath}");
                var ifcDict = TryLoadIfcDataFromPath(globalPath);
                if (ifcDict != null) return ifcDict;
            }

            RhinoApp.WriteLine("Loading IFC classes from default embedded resource.");
            return ReadIfcClassesFromResource("MyProject.Resources.Data.ifcClassListWithSubClass.csv");
        }

        // Public methods for resource loading (kept as fallback)
        public static void ReadCustomAttributeListsFromResource(string resourceName, out List<string> keys, out List<string> values)
        {
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
                    var result = ParseCustomCsvContent(reader.ReadToEnd());
                    keys = result.keys;
                    values = result.values;
                }
            }
        }

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