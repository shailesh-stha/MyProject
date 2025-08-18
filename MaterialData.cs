// MaterialData.cs
namespace MyProject
{
    /// <summary>
    /// Represents the data for a single material from the LCA reference file.
    /// </summary>
    public class MaterialData
    {
        public string MaterialName { get; set; }
        public double ReferenceQuantity { get; set; }
        public string ReferenceUnit { get; set; }
        public double Lca { get; set; }
    }
}