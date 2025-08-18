// CsvReader.cs
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Linq;

namespace MyProject
{
    /// <summary>
    /// A utility class to read data from an embedded CSV file.
    /// </summary>
    public static class CsvReader
    {
        /// <summary>
        /// Reads material LCA data from an embedded CSV resource file.
        /// The CSV is expected to have a header row and four columns: "MaterialName", "ReferenceQuantity", "ReferenceUnit", and "LCA".
        /// </summary>
        /// <param name="resourceName">The fully qualified name of the embedded resource.</param>
        /// <returns>A dictionary mapping material names (string) to their MaterialData objects.</returns>
        public static Dictionary<string, MaterialData> ReadMaterialLcaDataFromResource(string resourceName)
        {
            var materials = new Dictionary<string, MaterialData>();
            var assembly = Assembly.GetExecutingAssembly();

            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    Rhino.RhinoApp.WriteLine($"Error: Embedded resource '{resourceName}' not found.");
                    return materials;
                }

                using (StreamReader reader = new StreamReader(stream))
                {
                    // Skip the header row.
                    reader.ReadLine();

                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        var parts = line.Split(',');
                        if (parts.Length == 4)
                        {
                            var materialName = parts[0].Trim();
                            if (double.TryParse(parts[1], out double quantity) &&
                                double.TryParse(parts[3], out double lcaValue))
                            {
                                if (!string.IsNullOrWhiteSpace(materialName))
                                {
                                    var unit = parts[2].Trim();
                                    var materialData = new MaterialData
                                    {
                                        MaterialName = materialName,
                                        ReferenceQuantity = quantity,
                                        ReferenceUnit = unit,
                                        Lca = lcaValue
                                    };
                                    materials[materialName] = materialData;
                                }
                            }
                        }
                    }
                }
            }
            return materials;
        }
    }
}