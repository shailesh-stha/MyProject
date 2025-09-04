// PluginSettingsManager.cs
using Rhino.PlugIns;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MyProject
{
    /// <summary>
    /// A static helper class to manage reading and writing persistent plugin settings.
    /// This centralizes all settings keys and access logic.
    /// </summary>
    public static class PluginSettingsManager
    {
        // A prefix used for keys related to grid column layout.
        private const string GridColumnVisibilityPrefix = "GRID_COLUMN_VISIBILITY_";
        private const string GridColumnWidthPrefix = "GRID_COLUMN_WIDTH_";

        public static bool GetBool(string key, bool defaultValue) => MyProjectPlugin.Instance.Settings.GetBool(key, defaultValue);
        public static void SetBool(string key, bool value) => MyProjectPlugin.Instance.Settings.SetBool(key, value);
        public static string GetString(string key, string defaultValue) => MyProjectPlugin.Instance.Settings.GetString(key, defaultValue);
        public static void SetString(string key, string value) => MyProjectPlugin.Instance.Settings.SetString(key, value);
        public static int GetInt(string key, int defaultValue) => MyProjectPlugin.Instance.Settings.GetInteger(key, defaultValue);
        public static void SetInt(string key, int value) => MyProjectPlugin.Instance.Settings.SetInteger(key, value);

        /// <summary>
        /// Resets all saved settings for this plugin by manually overwriting them with their default values.
        /// </summary>
        public static void ResetToDefaults(IEnumerable<Eto.Forms.GridColumn> columns)
        {
            // --- MODIFICATION START ---
            // Manually set all settings to their hardcoded default values.
            // This overwrites the in-memory cache.

            // UI State & Preferences
            SetBool(SettingKeys.ShowAllObjects, true);
            SetBool(SettingKeys.ShowUnassigned, true);
            SetBool(SettingKeys.GroupByMaterial, false);
            SetBool(SettingKeys.GroupByClass, true);
            SetBool(SettingKeys.AggregateResults, false);
            SetBool(SettingKeys.DisplayIfcClass, false);
            SetBool(SettingKeys.DisplayMaterial, false);
            SetBool(SettingKeys.AssignDefOnCreate, false);
            SetBool(SettingKeys.AssignCustomOnCreate, false);

            // Viewport Display
            SetString(SettingKeys.IfcLeaderLength, "5");
            SetString(SettingKeys.IfcLeaderAngle, "35");
            SetString(SettingKeys.MaterialLeaderLength, "8");
            SetString(SettingKeys.MaterialLeaderAngle, "45");

            // Data Source Paths
            SetString(SettingKeys.MaterialCsvPath, string.Empty);
            SetString(SettingKeys.IfcClassCsvPath, string.Empty);
            SetString(SettingKeys.CustomCsvPath, string.Empty);

            // Expander States
            SetBool(SettingKeys.ExpanderUtilityButtons, true);
            SetBool(SettingKeys.ExpanderDefinition, true);
            SetBool(SettingKeys.ExpanderCustomDefinition, true);
            SetBool(SettingKeys.ExpanderAttributeGrid, true);
            SetBool(SettingKeys.ExpanderObjectAttributes, true);
            SetBool(SettingKeys.ExpanderLeaderOptions, true);
            SetBool(SettingKeys.ExpanderDocumentUserText, true);
            SetBool(SettingKeys.ExpanderAdvancedSettings, true);

            // Dynamically generated column layout settings
            if (columns != null)
            {
                foreach (var column in columns)
                {
                    var visibilityKey = $"{GridColumnVisibilityPrefix}{column.HeaderText}";
                    SetBool(visibilityKey, true); // Default to visible

                    var widthKey = $"{GridColumnWidthPrefix}{column.HeaderText}";
                    SetInt(widthKey, -1); // Default to auto-size
                }
            }

            MyProjectPlugin.Instance.SaveSettings(); // Force a write to disk
            Rhino.RhinoApp.WriteLine("Plugin settings have been reset to default.");
            // --- MODIFICATION END ---
        }

        public static void SaveColumnLayout(IEnumerable<Eto.Forms.GridColumn> columns)
        {
            if (columns == null) return;
            foreach (var column in columns)
            {
                var visibilityKey = $"{GridColumnVisibilityPrefix}{column.HeaderText}";
                SetBool(visibilityKey, column.Visible);

                var widthKey = $"{GridColumnWidthPrefix}{column.HeaderText}";
                SetInt(widthKey, column.Width);
            }
        }

        public static void LoadColumnLayout(IEnumerable<Eto.Forms.GridColumn> columns)
        {
            if (columns == null) return;
            foreach (var column in columns)
            {
                var visibilityKey = $"{GridColumnVisibilityPrefix}{column.HeaderText}";
                column.Visible = GetBool(visibilityKey, true);

                var widthKey = $"{GridColumnWidthPrefix}{column.HeaderText}";
                int savedWidth = GetInt(widthKey, -1);
                if (savedWidth > 0)
                {
                    column.Width = savedWidth;
                }
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
        public const string ExpanderUtilityButtons = "ExpanderUtilityButtons";
        public const string ExpanderDefinition = "ExpanderDefinition";
        public const string ExpanderCustomDefinition = "ExpanderCustomDefinition";
        public const string ExpanderAttributeGrid = "ExpanderAttributeGrid";
        public const string ExpanderObjectAttributes = "ExpanderObjectAttributes";
        public const string ExpanderLeaderOptions = "ExpanderLeaderOptions";
        public const string ExpanderDocumentUserText = "ExpanderDocumentUserText";
        public const string ExpanderAdvancedSettings = "ExpanderAdvancedSettings";
    }
}