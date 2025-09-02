// PluginSettingsManager.cs
using Rhino.PlugIns;
using System.Collections.Generic;
using System.Linq;

namespace MyProject
{
    /// <summary>
    /// A static helper class to manage reading and writing persistent plugin settings.
    /// This centralizes all settings keys and access logic.
    /// </summary>
    public static class PluginSettingsManager
    {
        // A prefix used for keys related to grid column visibility.
        private const string GridColumnVisibilityPrefix = "GRID_COLUMN_VISIBILITY_";

        /// <summary>
        /// Retrieves a boolean value from the plugin's persistent settings.
        /// </summary>
        /// <param name="key">The setting key.</param>
        /// <param name="defaultValue">The value to return if the key is not found.</param>
        /// <returns>The stored boolean value or the default value.</returns>
        public static bool GetBool(string key, bool defaultValue)
        {
            return MyProjectPlugin.Instance.Settings.GetBool(key, defaultValue);
        }

        /// <summary>
        /// Saves a boolean value to the plugin's persistent settings.
        /// </summary>
        /// <param name="key">The setting key.</param>
        /// <param name="value">The boolean value to save.</param>
        public static void SetBool(string key, bool value)
        {
            MyProjectPlugin.Instance.Settings.SetBool(key, value);
        }

        /// <summary>
        /// Retrieves a string value from the plugin's persistent settings.
        /// </summary>
        /// <param name="key">The setting key.</param>
        /// <param name="defaultValue">The value to return if the key is not found.</param>
        /// <returns>The stored string value or the default value.</returns>
        public static string GetString(string key, string defaultValue)
        {
            return MyProjectPlugin.Instance.Settings.GetString(key, defaultValue);
        }

        /// <summary>
        /// Saves a string value to the plugin's persistent settings.
        /// </summary>
        /// <param name="key">The setting key.</param>
        /// <param name="value">The string value to save.</param>
        public static void SetString(string key, string value)
        {
            MyProjectPlugin.Instance.Settings.SetString(key, value);
        }

        /// <summary>
        /// Saves the visibility state for a list of grid columns.
        /// </summary>
        /// <param name="columns">The collection of grid columns to save.</param>
        public static void SaveColumnVisibility(IEnumerable<Eto.Forms.GridColumn> columns)
        {
            if (columns == null) return;
            foreach (var column in columns)
            {
                var key = $"{GridColumnVisibilityPrefix}{column.HeaderText}";
                SetBool(key, column.Visible);
            }
        }

        /// <summary>
        /// Loads and applies the visibility state for a list of grid columns.
        /// </summary>
        /// <param name="columns">The collection of grid columns to update.</param>
        public static void LoadColumnVisibility(IEnumerable<Eto.Forms.GridColumn> columns)
        {
            if (columns == null) return;
            foreach (var column in columns)
            {
                var key = $"{GridColumnVisibilityPrefix}{column.HeaderText}";
                column.Visible = GetBool(key, true);
            }
        }
    }

    /// <summary>
    /// A static class to hold all the keys used for persistent settings.
    /// Using a central class for keys prevents errors from typos.
    /// </summary>
    public static class SettingKeys
    {
        public const string ShowAllObjects = "ShowAllObjects";
        public const string ShowUnassigned = "ShowUnassigned";
        public const string GroupByMaterial = "GroupByMaterial";
        public const string GroupByClass = "GroupByClass";
        public const string AggregateResults = "AggregateResults";
        public const string DisplayIfcClass = "DisplayIfcClass";
        public const string DisplayMaterial = "DisplayMaterial";
        public const string AssignDefOnCreate = "AssignDefOnCreate";
        public const string AssignCustomOnCreate = "AssignCustomOnCreate";
        public const string IfcLeaderLength = "IfcLeaderLength";
        public const string IfcLeaderAngle = "IfcLeaderAngle";
        public const string MaterialLeaderLength = "MaterialLeaderLength";
        public const string MaterialLeaderAngle = "MaterialLeaderAngle";
        public const string MaterialCsvPath = "MaterialCsvPath";
        public const string IfcClassCsvPath = "IfcClassCsvPath";
        public const string CustomCsvPath = "CustomCsvPath";
    }
}