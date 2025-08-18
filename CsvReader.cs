// CsvReader.cs
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace MyProject
{
    /// <summary>
    /// A utility class to read data from an embedded CSV file.
    /// </summary>
    public static class CsvReader
    {
        /// <summary>
        /// Reads a list of materials from an embedded CSV resource file.
        /// The CSV is expected to have one material name per line.
        /// </summary>
        /// <param name="resourceName">The fully qualified name of the embedded resource (e.g., "MyProject.Resources.materialList.csv").</param>
        /// <returns>A list of strings, where each string is a material name from the CSV.</returns>
        public static List<string> ReadMaterialsFromResource(string resourceName)
        {
            var materials = new List<string>();
            var assembly = Assembly.GetExecutingAssembly();

            // Use a stream to read the embedded resource file.
            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    // If the stream is null, it means the resource could not be found.
                    // This could be due to a typo in the resourceName or the file not being set as an "Embedded Resource".
                    Rhino.RhinoApp.WriteLine($"Error: Embedded resource '{resourceName}' not found.");
                    return materials; // Return an empty list.
                }

                using (StreamReader reader = new StreamReader(stream))
                {
                    string line;
                    // Read the file line by line until the end.
                    while ((line = reader.ReadLine()) != null)
                    {
                        // Trim any whitespace from the line and ensure it's not empty.
                        var materialName = line.Trim();
                        if (!string.IsNullOrWhiteSpace(materialName))
                        {
                            materials.Add(materialName);
                        }
                    }
                }
            }
            return materials;
        }
    }
}