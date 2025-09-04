// MyProjectPanel.cs

using Eto.Drawing;
using Eto.Forms;
using MyProject.Properties;
using Rhino;
using Rhino.Display;
using Rhino.DocObjects;
using Rhino.Geometry;
using Rhino.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace MyProject
{
    /// <summary>
    /// Represents the main Eto.Forms panel for the MyProject plugin.
    /// This panel displays object attributes, LCA data, and document information.
    /// </summary>
    [Guid("01D22D46-3C89-43A5-A92B-A1D6519A4244")]
    public class MyProjectPanel : Panel
    {
        #region Global Constants & Fields

        // --- Constants ---
        private const string PrimaryMaterialKey = "STR_MATERIAL_PR";
        private const string SecondaryMaterialPrefix = "STR_MATERIAL_SEC_";
        private const string IfcClassKey = "STR_IFC_CLASS";
        private const string QuantityMultiplierKey = "STR_MAT_PR_MULTIPLIER";
        private const string CustomAttributePrefix = "CUSTOM_";

        // --- Data Storage ---
        private Dictionary<string, List<string>> _ifcClasses = new Dictionary<string, List<string>>();
        private Dictionary<string, MaterialData> _materialLcaData;
        private List<string> _customAttributeKeys;
        private List<string> _customAttributeValues;

        private readonly ObservableCollection<ListItem> _allMaterialListItems = new ObservableCollection<ListItem>();
        private readonly FilterCollection<ListItem> _filteredMaterialItems;
        private readonly ObservableCollection<ListItem> _allIfcClassItems = new ObservableCollection<ListItem>();
        private readonly FilterCollection<ListItem> _filteredIfcClassItems;
        private readonly ObservableCollection<ListItem> _allIfcSubclassItems = new ObservableCollection<ListItem>();
        private readonly FilterCollection<ListItem> _filteredIfcSubclassItems;
        private readonly ObservableCollection<ListItem> _allCustomKeyItems = new ObservableCollection<ListItem>();
        private readonly FilterCollection<ListItem> _filteredCustomKeyItems;
        private readonly ObservableCollection<ListItem> _allCustomValueItems = new ObservableCollection<ListItem>();
        private readonly FilterCollection<ListItem> _filteredCustomValueItems;

        // --- Conduits ---
        private readonly AttributeDisplayConduit _ifcClassDisplayConduit;
        private readonly AttributeDisplayConduit _materialDisplayConduit;

        // --- State Flags ---
        private bool _isAssigningOnCreate = false;
        private bool _isSyncingSelection = false;
        private bool _needsRefresh = false;
        private bool _isMouseOverGrid = false;
        private string _originalObjectKey;

        #endregion

        #region Constructor & Main Panel Setup

        /// <summary>
        /// The panel constructor
        /// </summary>
        public MyProjectPanel()
        {
            _filteredMaterialItems = new FilterCollection<ListItem>(_allMaterialListItems);
            _filteredIfcClassItems = new FilterCollection<ListItem>(_allIfcClassItems);
            _filteredIfcSubclassItems = new FilterCollection<ListItem>(_allIfcSubclassItems);
            _filteredCustomKeyItems = new FilterCollection<ListItem>(_allCustomKeyItems);
            _filteredCustomValueItems = new FilterCollection<ListItem>(_allCustomValueItems);

            // Initialize conduits with settings
            _ifcClassDisplayConduit = new AttributeDisplayConduit
            {
                LeaderLength = int.Parse(PluginSettingsManager.GetString(SettingKeys.IfcLeaderLength, "5")),
                LeaderAngle = int.Parse(PluginSettingsManager.GetString(SettingKeys.IfcLeaderAngle, "35"))
            };
            _materialDisplayConduit = new AttributeDisplayConduit
            {
                LeaderLength = int.Parse(PluginSettingsManager.GetString(SettingKeys.MaterialLeaderLength, "8")),
                LeaderAngle = int.Parse(PluginSettingsManager.GetString(SettingKeys.MaterialLeaderAngle, "45"))
            };

            // Build the panel
            InitializeLayout();
            RegisterGlobalEventHandlers();
            LoadSettings();

            // Load initial data
            ReloadIfcClassList();
            ReloadMaterialList();
            ReloadCustomAttributeList();
            UpdatePanelData();
        }

        /// <summary>
        /// Creates the main layout and adds all the expandable sections.
        /// </summary>
        private void InitializeLayout()
        {
            Styles.Add<Label>("bold_label", label =>
            {
                label.Font = SystemFonts.Bold();
            });

            var mainLayout = new DynamicLayout { Spacing = new Size(6, 6), Padding = new Padding(10) };

            mainLayout.Add(CreateUtilityButtonsLayout());
            mainLayout.Add(CreateDefinitionLayout());
            mainLayout.Add(CreateCustomDefinitionLayout());
            mainLayout.Add(CreateAttributeGridLayout());
            mainLayout.Add(CreateObjectUserTextLayout());
            mainLayout.Add(CreateDisplayOptionsLayout());
            mainLayout.Add(CreateDocumentUserTextLayout());
            mainLayout.Add(CreateAdvancedSettingsLayout());
            mainLayout.Add(null, true);

            // --- Set visibility for all sections ---

            // These sections are set to be VISIBLE.
            _utilityButtonsExpander.Visible = true;
            _definitionExpander.Visible = true;
            _customDefinitionExpander.Visible = true;
            _attributeGridExpander.Visible = true;
            _objectAttributesExpander.Visible = true;
            _advancedSettingsExpander.Visible = true;

            // These sections are set to be HIDDEN.
            _viewportDisplayExpander.Visible = false;
            _documentUserTextExpander.Visible = false;


            Content = new Scrollable { Content = mainLayout, Border = BorderType.Bezel };
            MinimumSize = new Size(400, 450);
        }

        /// <summary>
        /// Registers event handlers for global Rhino events.
        /// Control-specific events are registered in their respective Create...Layout methods.
        /// </summary>
        private void RegisterGlobalEventHandlers()
        {
            RhinoDoc.SelectObjects += OnDocumentStateChanged;
            RhinoDoc.DeselectObjects += OnDocumentStateChanged;
            RhinoDoc.DeselectAllObjects += OnDocumentStateChanged;
            RhinoDoc.DeleteRhinoObject += OnDocumentStateChanged;
            RhinoDoc.ReplaceRhinoObject += OnDocumentStateChanged;
            RhinoDoc.UserStringChanged += OnDocumentStateChanged;
            RhinoDoc.EndOpenDocument += OnDocumentChanged;
            RhinoDoc.NewDocument += OnDocumentChanged;
            RhinoDoc.AddRhinoObject += OnObjectAdded;
            RhinoApp.Idle += OnRhinoIdle;
        }

        #endregion

        #region Section: Utility Buttons

        private Expander _utilityButtonsExpander;

        private Control CreateUtilityButtonsLayout()
        {
            // Create controls
            var structureIcon = BytesToEtoBitmap(Resources.btn_strLogo256, new Size(18, 18));
            var btnStructure = new Button
            {
                Image = structureIcon,
                ToolTip = "Visit str-ucture.com",
                MinimumSize = Size.Empty
            };

            var helloButton = new Button { Text = "Test" };
            var settingsButton = new Button { Text = "Settings", ToolTip = "Open plugin settings." };
            var resetButton = new Button { Text = "Reset to Default", ToolTip = "Resets all panel settings to their original defaults." };

            // Register events
            btnStructure.Click += (s, e) =>
            {
                var url = "http://www.str-ucture.com";
                try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
                catch (Exception ex) { RhinoApp.WriteLine($"Error opening website: {ex.Message}"); }
            };

            helloButton.Click += (s, e) =>
            {
                RhinoApp.WriteLine("hello");
            };

            settingsButton.Click += (s, e) =>
            {
                var dialog = new SettingsDialog();
                dialog.ShowModal(this);
            };

            resetButton.Click += (s, e) =>
            {
                var result = MessageBox.Show(
                    "Are you sure you want to reset all panel settings to their default values?\nThis will close and reopen the panel.",
                    "Confirm Reset",
                    MessageBoxButtons.YesNo,
                    MessageBoxType.Question
                );

                if (result == DialogResult.Yes)
                {
                    // Pass the grid columns to the reset method so it can reset dynamic keys.
                    PluginSettingsManager.ResetToDefaults(_userTextGridView.Columns);

                    // Close and reopen the panel to apply the now-defaulted settings.
                    var panelId = new Guid("01D22D46-3C89-43A5-A92B-A1D6519A4244");
                    Panels.ClosePanel(panelId);
                    Panels.OpenPanel(panelId);
                }
            };


            // Use a StackLayout to arrange buttons horizontally.
            var buttonLayout = new StackLayout
            {
                Orientation = Orientation.Horizontal,
                Spacing = 5,
                Items = { btnStructure, helloButton, settingsButton, resetButton }
            };

            // Assemble layout
            var layout = new DynamicLayout { Spacing = new Size(6, 6) };
            layout.AddRow(buttonLayout);

            _utilityButtonsExpander = new Expander { Header = new Label { Text = "Utility Buttons", Style = "bold_label" }, Content = layout };
            _utilityButtonsExpander.ExpandedChanged += (s, e) => PluginSettingsManager.SetBool(SettingKeys.ExpanderUtilityButtons, _utilityButtonsExpander.Expanded);

            return _utilityButtonsExpander;
        }

        private void OnExportToCsvClick(object sender, EventArgs e)
        {
            if (_userTextGridView.DataStore == null || !_userTextGridView.DataStore.Any())
            {
                RhinoApp.WriteLine("No data available to export.");
                return;
            }

            var data = _userTextGridView.DataStore.ToList();
            var dialog = new Eto.Forms.SaveFileDialog
            {
                Title = "Export Grid Data to CSV",
                FileName = $"LcaDataExport_{DateTime.Now:yyyyMMdd_HHmm}.csv",
                Filters = { new FileFilter("CSV Files (*.csv)", ".csv") }
            };

            if (dialog.ShowDialog(this) == DialogResult.Ok)
            {
                try
                {
                    var sb = new StringBuilder();
                    sb.AppendLine("ObjectIds,IfcClass,Material,Count,Quantity,QuantityUnit,QuantityMultiplier,QuantityTotal,LCA_kgCO2eq");

                    foreach (var entry in data)
                    {
                        var ids = string.Join(";", entry.ObjectIds);
                        var line = $"\"{ids}\",\"{entry.IfcClass}\",\"{entry.Value}\",{entry.Count},{entry.Quantity:F3},\"{entry.QuantityUnit}\",{entry.QuantityMultiplier:F3},{entry.QuantityTotal:F3},{entry.Lca:F3}";
                        sb.AppendLine(line);
                    }

                    File.WriteAllText(dialog.FileName, sb.ToString());
                    RhinoApp.WriteLine($"Data successfully exported to {dialog.FileName}");
                }
                catch (Exception ex)
                {
                    RhinoApp.WriteLine($"Error exporting data: {ex.Message}");
                }
            }
        }

        #endregion

        #region Section: Definition

        private Expander _definitionExpander;
        private readonly ComboBox _ifcDropdown = new ComboBox { Width = 135, AutoComplete = false };
        private readonly ComboBox _ifcSubclassDropdown = new ComboBox { Width = 90, AutoComplete = false };
        private readonly ComboBox _materialDropdown = new ComboBox { Width = 135, AutoComplete = false };
        private readonly ComboBox _materialTypeDropdown = new ComboBox { Width = 90 };
        private CheckBox _assignDefOnCreateCheckBox;

        private Control CreateDefinitionLayout()
        {
            // Create controls
            var selectIcon = BytesToEtoBitmap(Resources.btn_selectObjects256, new Size(18, 18));
            var selectIfcButton = new Button { Image = selectIcon, ToolTip = "Select objects matching the specified IfcClass.", MinimumSize = Size.Empty, ImagePosition = ButtonImagePosition.Left };
            var assignIfcButton = new Button { Image = BytesToEtoBitmap(Resources.btn_assignIfcClass256, new Size(18, 18)), ToolTip = "Assign selected IFC Class to object(s).", MinimumSize = Size.Empty, ImagePosition = ButtonImagePosition.Left };
            var removeIfcButton = new Button { Image = BytesToEtoBitmap(Resources.btn_removeIfcClass256, new Size(18, 18)), ToolTip = "Remove IFC Class from selected object(s).", MinimumSize = Size.Empty, ImagePosition = ButtonImagePosition.Left };
            var selectMaterialButton = new Button { Image = selectIcon, ToolTip = "Select objects matching the specified Material.", MinimumSize = Size.Empty, ImagePosition = ButtonImagePosition.Left };
            var assignMaterialButton = new Button { Image = BytesToEtoBitmap(Resources.btn_assignMaterial256, new Size(18, 18)), ToolTip = "Assign selected Material to object(s).", MinimumSize = Size.Empty, ImagePosition = ButtonImagePosition.Left };
            var removeMaterialButton = new Button { Image = BytesToEtoBitmap(Resources.btn_removeMaterial256, new Size(18, 18)), ToolTip = "Remove Material from selected object(s).", MinimumSize = Size.Empty, ImagePosition = ButtonImagePosition.Left };
            _assignDefOnCreateCheckBox = new CheckBox { Text = "Assign On Creation", ToolTip = "If checked, automatically assigns the selected IfcClass and Material to newly created objects." };

            // Setup static dropdown
            _materialTypeDropdown.Items.Add(new ListItem { Text = "Primary" });
            _materialTypeDropdown.Items.Add(new ListItem { Text = "Secondary" });
            _materialTypeDropdown.SelectedIndex = 0;

            // Generic ItemTextBinding for all filterable dropdowns
            var itemTextBinding = Binding.Property<ListItem, string>(li => li.Text);

            // IFC Class Dropdown
            _ifcDropdown.DataStore = _filteredIfcClassItems;
            _ifcDropdown.ItemTextBinding = itemTextBinding;
            _ifcDropdown.KeyUp += (s, e) => HandleComboBoxKeyUp(s as ComboBox, e, _filteredIfcClassItems);
            _ifcDropdown.MouseDown += (s, e) => UpdateFilterForComboBox(s as ComboBox, _filteredIfcClassItems);

            // IFC Subclass Dropdown
            _ifcSubclassDropdown.DataStore = _filteredIfcSubclassItems;
            _ifcSubclassDropdown.ItemTextBinding = itemTextBinding;
            _ifcSubclassDropdown.KeyUp += (s, e) => HandleComboBoxKeyUp(s as ComboBox, e, _filteredIfcSubclassItems);
            _ifcSubclassDropdown.MouseDown += (s, e) => UpdateFilterForComboBox(s as ComboBox, _filteredIfcSubclassItems);

            // Material Dropdown
            _materialDropdown.DataStore = _filteredMaterialItems;
            _materialDropdown.ItemTextBinding = itemTextBinding;
            _materialDropdown.KeyUp += (s, e) => HandleComboBoxKeyUp(s as ComboBox, e, _filteredMaterialItems);
            _materialDropdown.MouseDown += (s, e) => UpdateFilterForComboBox(s as ComboBox, _filteredMaterialItems);

            // Register other events
            selectIfcButton.Click += OnSelectByIfcClassClick;
            assignIfcButton.Click += OnAssignIfcClassClick;
            removeIfcButton.Click += (s, e) => RemoveUserString(IfcClassKey, "IfcClass");
            selectMaterialButton.Click += OnSelectByMaterialClick;
            assignMaterialButton.Click += OnAssignMaterialClick;
            removeMaterialButton.Click += OnRemoveMaterialClick;
            _ifcDropdown.SelectedValueChanged += OnPrimaryIfcClassChanged;
            _assignDefOnCreateCheckBox.CheckedChanged += (s, e) => PluginSettingsManager.SetBool(SettingKeys.AssignDefOnCreate, _assignDefOnCreateCheckBox.Checked ?? false);

            // Assemble layout
            var layout = new DynamicLayout { Spacing = new Size(6, 6) };
            var ifcClassLayout = new StackLayout { Orientation = Orientation.Horizontal, Spacing = 5, Items = { _ifcDropdown, _ifcSubclassDropdown } };
            var ifcButtonsLayout = new StackLayout { Orientation = Orientation.Horizontal, Spacing = 5, Items = { assignIfcButton, removeIfcButton, selectIfcButton } };
            layout.AddRow(new Label { Text = "IfcClass:", ToolTip = IfcClassKey }, ifcClassLayout, ifcButtonsLayout);

            var materialDropdownsLayout = new StackLayout { Orientation = Orientation.Horizontal, Spacing = 5, Items = { _materialDropdown, _materialTypeDropdown } };
            var materialButtonsLayout = new StackLayout { Orientation = Orientation.Horizontal, Spacing = 5, Items = { assignMaterialButton, removeMaterialButton, selectMaterialButton } };
            layout.AddRow(new Label { Text = "Material:", ToolTip = $"{PrimaryMaterialKey} / {SecondaryMaterialPrefix}N" }, materialDropdownsLayout, materialButtonsLayout);
            layout.AddRow(null, _assignDefOnCreateCheckBox);

            _definitionExpander = new Expander { Header = new Label { Text = "Primary Attribute Assignment", Style = "bold_label" }, Content = layout };
            _definitionExpander.ExpandedChanged += (s, e) => PluginSettingsManager.SetBool(SettingKeys.ExpanderDefinition, _definitionExpander.Expanded);

            return _definitionExpander;
        }

        /// <summary>
        /// A single, reusable KeyUp event handler for all filterable ComboBoxes.
        /// </summary>
        private void HandleComboBoxKeyUp(ComboBox comboBox, KeyEventArgs e, FilterCollection<ListItem> filterCollection)
        {
            if (comboBox == null || e.Key == Keys.Up || e.Key == Keys.Down || e.Key == Keys.Enter)
                return;

            UpdateFilterForComboBox(comboBox, filterCollection);
        }

        /// <summary>
        /// A single, reusable method to apply a text filter to a ComboBox's FilterCollection.
        /// </summary>
        private void UpdateFilterForComboBox(ComboBox comboBox, FilterCollection<ListItem> filterCollection)
        {
            if (comboBox == null || filterCollection == null) return;

            var searchString = comboBox.Text;

            if (string.IsNullOrWhiteSpace(searchString))
            {
                filterCollection.Filter = null;
            }
            else
            {
                filterCollection.Filter = item =>
                {
                    if (item?.Text == null) return false;
                    // Perform a case-insensitive "contains" search.
                    return item.Text.IndexOf(searchString, StringComparison.OrdinalIgnoreCase) >= 0;
                };
            }
        }

        private void OnSelectByIfcClassClick(object sender, EventArgs e)
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;

            string primaryClass = (_ifcDropdown.SelectedValue as ListItem)?.Text;
            if (string.IsNullOrWhiteSpace(primaryClass))
            {
                RhinoApp.WriteLine("Please select an IfcClass to search for.");
                return;
            }

            string secondaryClass = (_ifcSubclassDropdown.SelectedValue as ListItem)?.Text;
            string ifcTarget = !string.IsNullOrWhiteSpace(secondaryClass) ? $"{primaryClass}, {secondaryClass}" : primaryClass;

            var idsToSelect = doc.Objects
                .Where(obj => obj.Attributes.GetUserString(IfcClassKey) == ifcTarget)
                .Select(obj => obj.Id)
                .ToList();

            doc.Objects.UnselectAll();
            if (idsToSelect.Any())
            {
                doc.Objects.Select(idsToSelect, true);
                RhinoApp.WriteLine($"{idsToSelect.Count} object(s) selected with IfcClass '{ifcTarget}'.");
            }
            else
            {
                RhinoApp.WriteLine($"No objects found with IfcClass '{ifcTarget}'.");
            }
            doc.Views.Redraw();
        }

        private void OnAssignIfcClassClick(object sender, EventArgs e)
        {
            string primaryClass = (_ifcDropdown.SelectedValue as ListItem)?.Text;
            string secondaryClass = (_ifcSubclassDropdown.SelectedValue as ListItem)?.Text;

            if (string.IsNullOrWhiteSpace(primaryClass) && string.IsNullOrWhiteSpace(secondaryClass))
            {
                RhinoApp.WriteLine("Please select an IFC class from the dropdowns first.");
                return;
            }

            if (!string.IsNullOrWhiteSpace(secondaryClass) && string.IsNullOrWhiteSpace(primaryClass))
            {
                RhinoApp.WriteLine("A primary IFC class must be selected when specifying a secondary class.");
                return;
            }

            string combinedClass = !string.IsNullOrWhiteSpace(secondaryClass) ? $"{primaryClass}, {secondaryClass}" : primaryClass;
            AssignUserString(IfcClassKey, "IfcClass", combinedClass);
        }

        private void OnSelectByMaterialClick(object sender, EventArgs e)
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;

            string materialTarget = (_materialDropdown.SelectedValue as ListItem)?.Text;
            if (string.IsNullOrWhiteSpace(materialTarget))
            {
                RhinoApp.WriteLine("Please select a Material to search for.");
                return;
            }

            var idsToSelect = doc.Objects
                .Where(obj =>
                {
                    var userStrings = obj.Attributes.GetUserStrings();
                    if (userStrings.Count == 0) return false;
                    // Check primary material
                    if (userStrings.Get(PrimaryMaterialKey) == materialTarget) return true;
                    // Check all secondary materials
                    return userStrings.AllKeys
                        .Where(k => k.StartsWith(SecondaryMaterialPrefix))
                        .Any(k => userStrings.Get(k) == materialTarget);
                })
                .Select(obj => obj.Id)
                .ToList();

            doc.Objects.UnselectAll();
            if (idsToSelect.Any())
            {
                doc.Objects.Select(idsToSelect, true);
                RhinoApp.WriteLine($"{idsToSelect.Count} object(s) selected with Material '{materialTarget}'.");
            }
            else
            {
                RhinoApp.WriteLine($"No objects found with Material '{materialTarget}'.");
            }
            doc.Views.Redraw();
        }

        private void OnAssignMaterialClick(object sender, EventArgs e)
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;

            var selectedObjects = doc.Objects.GetSelectedObjects(false, false).ToList();
            if (!selectedObjects.Any())
            {
                RhinoApp.WriteLine("Please select one or more objects to assign material to.");
                return;
            }

            string materialValue = (_materialDropdown.SelectedValue as ListItem)?.Text;
            if (string.IsNullOrWhiteSpace(materialValue))
            {
                RhinoApp.WriteLine("Please select a material from the dropdown.");
                return;
            }

            string materialType = (_materialTypeDropdown.SelectedValue as ListItem)?.Text;
            int updatedCount = 0;

            foreach (var rhinoObject in selectedObjects)
            {
                var newAttributes = rhinoObject.Attributes.Duplicate();
                bool attributeSet = false;

                if (materialType == "Primary")
                {
                    attributeSet = newAttributes.SetUserString(PrimaryMaterialKey, materialValue);
                }
                else // Secondary
                {
                    var userStrings = newAttributes.GetUserStrings();
                    int nextIndex = 1;
                    if (userStrings != null)
                    {
                        var secondaryKeys = userStrings.AllKeys
                                                       .Where(k => k.StartsWith(SecondaryMaterialPrefix))
                                                       .Select(k => int.TryParse(k.Replace(SecondaryMaterialPrefix, ""), out int index) ? index : 0)
                                                       .ToList();
                        if (secondaryKeys.Any())
                        {
                            nextIndex = secondaryKeys.Max() + 1;
                        }
                    }
                    var newKey = $"{SecondaryMaterialPrefix}{nextIndex}";
                    attributeSet = newAttributes.SetUserString(newKey, materialValue);
                }

                if (attributeSet && doc.Objects.ModifyAttributes(rhinoObject.Id, newAttributes, false))
                {
                    updatedCount++;
                }
            }

            if (updatedCount > 0)
            {
                RhinoApp.WriteLine($"Successfully assigned material '{materialValue}' ({materialType}) to {updatedCount} object(s).");
                doc.Views.Redraw();
            }
        }

        private void OnRemoveMaterialClick(object sender, EventArgs e)
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;

            var selectedObjects = doc.Objects.GetSelectedObjects(false, false).ToList();
            if (!selectedObjects.Any())
            {
                RhinoApp.WriteLine("Please select one or more objects to remove materials from.");
                return;
            }

            int updatedCount = 0;
            foreach (var rhinoObject in selectedObjects)
            {
                var newAttributes = rhinoObject.Attributes.Duplicate();
                var userStrings = newAttributes.GetUserStrings();
                if (userStrings == null || userStrings.Count == 0) continue;

                var secondaryKeys = userStrings.AllKeys
                    .Where(k => k.StartsWith(SecondaryMaterialPrefix))
                    .Select(k => new { Key = k, Index = int.TryParse(k.Replace(SecondaryMaterialPrefix, ""), out int index) ? index : 0 })
                    .OrderByDescending(k => k.Index)
                    .ToList();

                bool removed = false;
                if (secondaryKeys.Any())
                {
                    // Remove the secondary material with the highest index
                    var keyToRemove = secondaryKeys.First().Key;
                    newAttributes.DeleteUserString(keyToRemove);
                    removed = true;
                    RhinoApp.WriteLine($"Removing secondary material '{userStrings.Get(keyToRemove)}' from object {rhinoObject.Id}.");
                }
                else if (userStrings.Get(PrimaryMaterialKey) != null)
                {
                    // No secondary materials left, remove the primary
                    newAttributes.DeleteUserString(PrimaryMaterialKey);
                    removed = true;
                    RhinoApp.WriteLine($"Removing primary material '{userStrings.Get(PrimaryMaterialKey)}' from object {rhinoObject.Id}.");
                }

                if (removed && doc.Objects.ModifyAttributes(rhinoObject.Id, newAttributes, false))
                {
                    updatedCount++;
                }
            }

            if (updatedCount > 0)
            {
                RhinoApp.WriteLine($"Successfully processed material removal for {updatedCount} object(s).");
                doc.Views.Redraw();
            }
            else
            {
                RhinoApp.WriteLine("Selected objects have no materials to remove.");
            }
        }

        private void OnPrimaryIfcClassChanged(object sender, EventArgs e) => UpdateSubclassDropdown();

        #endregion

        #region Section: Custom Definition

        private Expander _customDefinitionExpander;
        private ComboBox _customKeyDropdown = new ComboBox { Width = 135, AutoComplete = false };
        private ComboBox _customValueDropdown = new ComboBox { Width = 90, AutoComplete = false };
        private CheckBox _assignCustomOnCreateCheckBox;

        private Control CreateCustomDefinitionLayout()
        {
            // Create controls
            var selectIcon = BytesToEtoBitmap(Resources.btn_selectObjects256, new Size(18, 18));
            var selectCustomButton = new Button { Image = selectIcon, ToolTip = "Select objects matching the specified custom key and/or value.", MinimumSize = Size.Empty };
            var assignButton = new Button { Image = BytesToEtoBitmap(Resources.btn_assignCustom256, new Size(18, 18)), ToolTip = "Assign selected Custom Key/Value to object(s).", MinimumSize = Size.Empty, ImagePosition = ButtonImagePosition.Left };
            var removeButton = new Button { Image = BytesToEtoBitmap(Resources.btn_removeCustom256, new Size(18, 18)), ToolTip = "Remove Custom Key/Value from selected object(s).", MinimumSize = Size.Empty, ImagePosition = ButtonImagePosition.Left };
            _assignCustomOnCreateCheckBox = new CheckBox { Text = "Assign On Creation", ToolTip = "If checked, automatically assigns the selected custom attribute to newly created objects." };

            var itemTextBinding = Binding.Property<ListItem, string>(li => li.Text);

            _customKeyDropdown.DataStore = _filteredCustomKeyItems;
            _customKeyDropdown.ItemTextBinding = itemTextBinding;
            _customKeyDropdown.KeyUp += (s, e) => HandleComboBoxKeyUp(s as ComboBox, e, _filteredCustomKeyItems);
            _customKeyDropdown.MouseDown += (s, e) => UpdateFilterForComboBox(s as ComboBox, _filteredCustomKeyItems);

            _customValueDropdown.DataStore = _filteredCustomValueItems;
            _customValueDropdown.ItemTextBinding = itemTextBinding;
            _customValueDropdown.KeyUp += (s, e) => HandleComboBoxKeyUp(s as ComboBox, e, _filteredCustomValueItems);
            _customValueDropdown.MouseDown += (s, e) => UpdateFilterForComboBox(s as ComboBox, _filteredCustomValueItems);

            // Register other events
            selectCustomButton.Click += OnSelectByCustomAttributeClick;
            assignButton.Click += OnAssignCustomAttributeClick;
            removeButton.Click += OnRemoveCustomAttributeClick;
            _assignCustomOnCreateCheckBox.CheckedChanged += (s, e) => PluginSettingsManager.SetBool(SettingKeys.AssignCustomOnCreate, _assignCustomOnCreateCheckBox.Checked ?? false);

            // Assemble layout
            var layout = new DynamicLayout { Spacing = new Size(6, 6) };
            var dropdownLayout = new StackLayout { Orientation = Orientation.Horizontal, Spacing = 5, Items = { _customKeyDropdown, _customValueDropdown } };
            var buttonsLayout = new StackLayout { Orientation = Orientation.Horizontal, Spacing = 5, Items = { assignButton, removeButton, selectCustomButton } };
            layout.AddRow(new Label { Text = "Custom:", ToolTip = "CUSTOM KEY." }, dropdownLayout, buttonsLayout);
            layout.AddRow(null, _assignCustomOnCreateCheckBox);

            _customDefinitionExpander = new Expander { Header = new Label { Text = "Custom Attribute Assignment", Style = "bold_label" }, Content = layout };
            _customDefinitionExpander.ExpandedChanged += (s, e) => PluginSettingsManager.SetBool(SettingKeys.ExpanderCustomDefinition, _customDefinitionExpander.Expanded);

            return _customDefinitionExpander;
        }

        private void OnSelectByCustomAttributeClick(object sender, EventArgs e)
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;

            string keyTarget = (_customKeyDropdown.SelectedValue as ListItem)?.Text;
            string valueTarget = (_customValueDropdown.SelectedValue as ListItem)?.Text;

            if (string.IsNullOrWhiteSpace(keyTarget) && string.IsNullOrWhiteSpace(valueTarget))
            {
                RhinoApp.WriteLine("Please select a custom key or value to search for.");
                return;
            }

            var idsToSelect = new List<Guid>();
            var allObjects = doc.Objects.GetObjectList(ObjectType.AnyObject);
            string formattedKeyTarget = string.IsNullOrWhiteSpace(keyTarget) ? null : $"{CustomAttributePrefix}{keyTarget}";

            foreach (var obj in allObjects)
            {
                var userStrings = obj.Attributes.GetUserStrings();
                if (userStrings == null || userStrings.Count == 0) continue;

                bool match = false;

                if (!string.IsNullOrWhiteSpace(formattedKeyTarget) && !string.IsNullOrWhiteSpace(valueTarget))
                {
                    match = userStrings.Get(formattedKeyTarget) == valueTarget;
                }
                else if (!string.IsNullOrWhiteSpace(formattedKeyTarget))
                {
                    match = userStrings.Get(formattedKeyTarget) != null;
                }
                else if (!string.IsNullOrWhiteSpace(valueTarget))
                {
                    match = userStrings.AllKeys
                                       .Where(k => k.StartsWith(CustomAttributePrefix))
                                       .Any(k => userStrings.Get(k) == valueTarget);
                }

                if (match)
                {
                    idsToSelect.Add(obj.Id);
                }
            }

            doc.Objects.UnselectAll();
            if (idsToSelect.Any())
            {
                doc.Objects.Select(idsToSelect, true);
                RhinoApp.WriteLine($"{idsToSelect.Count} matching object(s) selected.");
            }
            else
            {
                RhinoApp.WriteLine("No objects found with the specified custom criteria.");
            }
            doc.Views.Redraw();
        }

        private void OnAssignCustomAttributeClick(object sender, EventArgs e)
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;

            string key = (_customKeyDropdown.SelectedValue as ListItem)?.Text;
            string value = (_customValueDropdown.SelectedValue as ListItem)?.Text;

            if (string.IsNullOrWhiteSpace(key) || string.IsNullOrWhiteSpace(value))
            {
                RhinoApp.WriteLine("Please select a key and a value for the custom attribute.");
                return;
            }

            var selectedObjects = doc.Objects.GetSelectedObjects(false, false).ToList();
            if (!selectedObjects.Any())
            {
                RhinoApp.WriteLine("Please select one or more objects to assign the attribute to.");
                return;
            }

            string formattedKey = $"{CustomAttributePrefix}{key}";
            int updatedCount = 0;

            foreach (var rhinoObject in selectedObjects)
            {
                var newAttributes = rhinoObject.Attributes.Duplicate();
                var keysToRemove = newAttributes.GetUserStrings().AllKeys
                                     .Where(k => k.StartsWith(CustomAttributePrefix))
                                     .ToList();
                foreach (var keyToRemove in keysToRemove)
                {
                    newAttributes.DeleteUserString(keyToRemove);
                }

                if (newAttributes.SetUserString(formattedKey, value) && doc.Objects.ModifyAttributes(rhinoObject.Id, newAttributes, false))
                {
                    updatedCount++;
                }
            }
            if (updatedCount > 0)
            {
                RhinoApp.WriteLine($"Successfully assigned custom attribute to {updatedCount} object(s).");
                doc.Views.Redraw();
            }
        }

        private void OnRemoveCustomAttributeClick(object sender, EventArgs e)
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;

            var selectedObjects = doc.Objects.GetSelectedObjects(false, false).ToList();
            if (!selectedObjects.Any())
            {
                RhinoApp.WriteLine("Please select one or more objects to remove the custom attribute from.");
                return;
            }

            int updatedCount = 0;
            foreach (var rhinoObject in selectedObjects)
            {
                bool objectModified = false;
                var newAttributes = rhinoObject.Attributes.Duplicate();
                var keysToRemove = newAttributes.GetUserStrings().AllKeys
                                     .Where(k => k.StartsWith(CustomAttributePrefix))
                                     .ToList();

                if (keysToRemove.Any())
                {
                    foreach (var keyToRemove in keysToRemove)
                    {
                        newAttributes.DeleteUserString(keyToRemove);
                    }
                    if (doc.Objects.ModifyAttributes(rhinoObject.Id, newAttributes, false))
                    {
                        objectModified = true;
                    }
                }
                if (objectModified) updatedCount++;
            }

            if (updatedCount > 0)
            {
                RhinoApp.WriteLine($"Successfully removed custom attribute from {updatedCount} object(s).");
                doc.Views.Redraw();
            }
            else
            {
                RhinoApp.WriteLine("Selected objects do not have a custom attribute to remove.");
            }
        }

        #endregion

        #region Section: Attribute Grid (LCA Calculation)

        private Expander _attributeGridExpander;
        private readonly GridView<UserTextEntry> _userTextGridView = new GridView<UserTextEntry> { ShowHeader = true, AllowMultipleSelection = true, Width = 450 };
        private readonly GridColumn _qtyMultiplierColumn = new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.QuantityMultiplier)) { TextAlignment = TextAlignment.Right }, HeaderText = "Qty. Multiplier", Editable = true, Width = 90 };
        private readonly Label _totalLcaLabel = new Label { Text = "Total LCA: 0.00" };
        private readonly CheckBox _showAllObjectsCheckBox = new CheckBox { Text = "Show All Objects" };
        private readonly CheckBox _showUnassignedCheckBox = new CheckBox { Text = "Show/Hide Unassigned" };
        private readonly CheckBox _groupByMaterialCheckBox = new CheckBox { Text = "Group by Material" };
        private readonly CheckBox _groupByClassCheckBox = new CheckBox { Text = "Group by Class" };
        private readonly CheckBox _aggregateCheckBox = new CheckBox { Text = "Aggregate" };

        private Control CreateAttributeGridLayout()
        {
            _totalLcaLabel.Font = SystemFonts.Bold();

            // Setup grid columns if they don't exist
            if (_userTextGridView.Columns.Count == 0)
            {
                _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.SerialNumber)), HeaderText = "SN", Editable = false, Width = 30 });
                _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.GroupName)), HeaderText = "Group/Block", Editable = false, Width = 120 });
                _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.IfcClass)), HeaderText = "IfcClass", Editable = false, Width = 150 });
                _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.Value)), HeaderText = "Material (Name)", Editable = false, Width = 150 });
                _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.DisplayReference)) { TextAlignment = TextAlignment.Right }, HeaderText = "Ref. Qty.", Editable = false, Width = 70 });
                _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell { Binding = Binding.Property<UserTextEntry, string>(entry => $"{entry.ReferenceLca:0.00}"), TextAlignment = TextAlignment.Right }, HeaderText = "Ref. LCA", Editable = false, Width = 70 });
                _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.Count)) { TextAlignment = TextAlignment.Right }, HeaderText = "Count", Editable = false, Width = 50, Visible = false });
                _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.DisplayQuantity)) { TextAlignment = TextAlignment.Right }, HeaderText = "Qty. (Rh)", Editable = false, Width = 75 });
                _userTextGridView.Columns.Add(_qtyMultiplierColumn);
                _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell { Binding = Binding.Property<UserTextEntry, string>(entry => $"{entry.QuantityTotal:0.00}"), TextAlignment = TextAlignment.Right }, HeaderText = "Qty. Total", Editable = false, Width = 75 });
                _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell { Binding = Binding.Property<UserTextEntry, string>(entry => $"{entry.Lca:0.00}"), TextAlignment = TextAlignment.Right }, HeaderText = "LCA (kgCO2 eq)", Editable = false, Width = 100 });
            }
            _qtyMultiplierColumn.Visible = !_aggregateCheckBox.Checked ?? true;

            // Create controls
            var columnsViewButton = new Button { Image = BytesToEtoBitmap(Resources.btn_columnsView256, new Size(18, 18)), MinimumSize = Size.Empty, ImagePosition = ButtonImagePosition.Left };
            var selectUnassignedButton = new Button { Image = BytesToEtoBitmap(Resources.btn_selectUnassigned256, new Size(18, 18)), ToolTip = "Select objects with no material assigned.", MinimumSize = Size.Empty };

            var exportButton = new Button { Text = "Exp.", ToolTip = "Export the LCA table data to a CSV file." };
            exportButton.Click += OnExportToCsvClick;

            // Register events
            columnsViewButton.Click += (s, e) =>
            {
                new ColumnVisibilityDialog(_userTextGridView).ShowModal(this);
                PluginSettingsManager.SaveColumnLayout(_userTextGridView.Columns);
            };
            selectUnassignedButton.Click += OnSelectUnassignedClick;
            _showAllObjectsCheckBox.CheckedChanged += (s, e) => { PluginSettingsManager.SetBool(SettingKeys.ShowAllObjects, _showAllObjectsCheckBox.Checked ?? false); UpdatePanelData(); };
            _showUnassignedCheckBox.CheckedChanged += (s, e) => { PluginSettingsManager.SetBool(SettingKeys.ShowUnassigned, _showUnassignedCheckBox.Checked ?? false); UpdatePanelData(); };
            _aggregateCheckBox.CheckedChanged += (s, e) => { PluginSettingsManager.SetBool(SettingKeys.AggregateResults, _aggregateCheckBox.Checked ?? false); OnAggregateChanged(s, e); };
            _groupByMaterialCheckBox.CheckedChanged += (s, e) => { PluginSettingsManager.SetBool(SettingKeys.GroupByMaterial, _groupByMaterialCheckBox.Checked ?? false); OnGroupByMaterialChanged(s, e); };
            _groupByClassCheckBox.CheckedChanged += (s, e) => { PluginSettingsManager.SetBool(SettingKeys.GroupByClass, _groupByClassCheckBox.Checked ?? false); OnGroupByClassChanged(s, e); };
            _userTextGridView.SelectionChanged += OnGridSelectionChanged;
            _userTextGridView.CellEdited += OnGridCellEdited;
            _userTextGridView.MouseEnter += (s, e) => _isMouseOverGrid = true;
            _userTextGridView.MouseLeave += (s, e) => _isMouseOverGrid = false;

            // Assemble layout
            var layout = new DynamicLayout { Spacing = new Size(6, 6) };
            var viewOptionsLayout = new StackLayout { Orientation = Orientation.Horizontal, Spacing = 10, Items = { _showAllObjectsCheckBox, _showUnassignedCheckBox } };

            var groupingOptionsLayout = new StackLayout { Orientation = Orientation.Horizontal, Spacing = 10, Items = { _groupByClassCheckBox, _groupByMaterialCheckBox, _aggregateCheckBox, null, selectUnassignedButton, columnsViewButton, exportButton } };

            layout.AddRow(viewOptionsLayout);
            layout.AddRow(groupingOptionsLayout);
            layout.AddRow(new Scrollable { Content = _userTextGridView, Width = 450, Height = 200 });
            layout.AddRow(_totalLcaLabel);

            _attributeGridExpander = new Expander { Header = new Label { Text = "LCA Calculation Table", Style = "bold_label" }, Content = layout };
            _attributeGridExpander.ExpandedChanged += (s, e) => PluginSettingsManager.SetBool(SettingKeys.ExpanderAttributeGrid, _attributeGridExpander.Expanded);

            return _attributeGridExpander;
        }

        private void OnAggregateChanged(object sender, EventArgs e) => UpdatePanelData();

        private void OnGroupByClassChanged(object sender, EventArgs e)
        {
            if (_groupByClassCheckBox.Checked == true) _groupByMaterialCheckBox.Checked = false;
            UpdatePanelData();
        }

        private void OnGroupByMaterialChanged(object sender, EventArgs e)
        {
            if (_groupByMaterialCheckBox.Checked == true) _groupByClassCheckBox.Checked = false;
            ReorderGridColumns(_groupByMaterialCheckBox.Checked ?? false);
            UpdatePanelData();
        }

        private void OnSelectUnassignedClick(object sender, EventArgs e)
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;

            var idsToSelect = doc.Objects
              .Where(o => o.IsSelectable(true, false, false, true) &&
                    o.Attributes.Visible &&
                    doc.Layers[o.Attributes.LayerIndex].IsVisible &&
                    string.IsNullOrWhiteSpace(o.Attributes.GetUserString(PrimaryMaterialKey)))
              .Select(o => o.Id)
              .ToList();

            doc.Objects.UnselectAll();
            if (idsToSelect.Any())
            {
                doc.Objects.Select(idsToSelect, true);
                RhinoApp.WriteLine($"{idsToSelect.Count} unassigned object(s) selected.");
            }
            else
            {
                RhinoApp.WriteLine("No unassigned objects found.");
            }
            doc.Views.Redraw();
        }

        private void OnGridCellEdited(object sender, GridViewCellEventArgs e)
        {
            if (e.Item is UserTextEntry entry && entry.ObjectIds.Any())
            {
                var doc = RhinoDoc.ActiveDoc;
                if (doc == null) return;

                var objectId = entry.ObjectIds.First();
                var rhinoObject = doc.Objects.FindId(objectId);
                if (rhinoObject != null)
                {
                    var newAttributes = rhinoObject.Attributes.Duplicate();
                    var multiplierValue = entry.QuantityMultiplier.ToString(CultureInfo.InvariantCulture);
                    newAttributes.SetUserString(QuantityMultiplierKey, multiplierValue);
                    doc.Objects.ModifyAttributes(rhinoObject.Id, newAttributes, false);
                }
                UpdatePanelData();
            }
        }

        private void OnGridSelectionChanged(object sender, EventArgs e)
        {
            if (!_isMouseOverGrid || _isSyncingSelection) return;

            _isSyncingSelection = true;
            var doc = RhinoDoc.ActiveDoc;
            if (doc != null)
            {
                var gridSelectionIds = _userTextGridView.SelectedItems.SelectMany(entry => entry.ObjectIds).ToHashSet();
                var rhinoSelectionIds = doc.Objects.GetSelectedObjects(false, false).Select(o => o.Id).ToHashSet();
                var idsToSelect = gridSelectionIds.Except(rhinoSelectionIds).ToList();
                if (idsToSelect.Any()) doc.Objects.Select(idsToSelect, true);
                var idsToDeselect = rhinoSelectionIds.Except(gridSelectionIds).ToList();
                if (idsToDeselect.Any())
                {
                    foreach (var id in idsToDeselect) doc.Objects.Select(id, false);
                }
                doc.Views.Redraw();
            }
            _isSyncingSelection = false;
        }

        #endregion

        #region Section: Object Attributes

        private Expander _objectAttributesExpander;
        private Label _selectionCountLabel;
        private readonly GridView<ObjectUserTextEntry> _objectUserTextGridView = new GridView<ObjectUserTextEntry> { ShowHeader = true, Height = 100, Width = 450 };
        private readonly GridColumn _objectKeyColumn = new GridColumn { HeaderText = "Key", DataCell = new TextBoxCell(nameof(ObjectUserTextEntry.Key)), Editable = false, Width = 150 };
        private readonly GridColumn _objectValueColumn = new GridColumn { HeaderText = "Value", DataCell = new TextBoxCell(nameof(ObjectUserTextEntry.Value)), Editable = false, Width = 250 };

        private Control CreateObjectUserTextLayout()
        {
            if (_objectUserTextGridView.Columns.Count == 0)
            {
                _objectUserTextGridView.Columns.Add(new GridColumn
                {
                    HeaderText = "SN",
                    DataCell = new TextBoxCell(nameof(ObjectUserTextEntry.SerialNumber)),
                    Editable = false,
                    Width = 30
                });
                _objectUserTextGridView.Columns.Add(_objectKeyColumn);
                _objectUserTextGridView.Columns.Add(_objectValueColumn);
            }

            // Register events
            _objectUserTextGridView.CellEditing += OnObjectGridCellEditing;
            _objectUserTextGridView.CellEdited += OnObjectGridCellEdited;

            _selectionCountLabel = new Label { Text = "No Object Selected" };

            var layout = new DynamicLayout { Spacing = new Size(6, 6) };
            layout.AddRow(_selectionCountLabel);
            layout.AddRow(new Scrollable { Content = _objectUserTextGridView, Width = 450, Height = 120 });

            var headerLabel = new Label { Text = "Object Inspector", Style = "bold_label" };
            _objectAttributesExpander = new Expander { Header = headerLabel, Content = layout };

            _objectAttributesExpander.ExpandedChanged += (s, e) => PluginSettingsManager.SetBool(SettingKeys.ExpanderObjectAttributes, _objectAttributesExpander.Expanded);

            return _objectAttributesExpander;
        }

        private void OnObjectGridCellEditing(object sender, GridViewCellEventArgs e)
        {
            if (e.Item is ObjectUserTextEntry entry)
            {
                _originalObjectKey = entry.Key;
            }
            else
            {
                _originalObjectKey = null;
            }
        }

        private void OnObjectGridCellEdited(object sender, GridViewCellEventArgs e)
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null || string.IsNullOrEmpty(_originalObjectKey)) return;

            var selectedObjects = doc.Objects.GetSelectedObjects(false, false).ToList();
            if (selectedObjects.Count != 1) return;

            var rhinoObject = selectedObjects[0];
            if (rhinoObject == null) return;

            if (e.Item is ObjectUserTextEntry entry)
            {
                var newAttributes = rhinoObject.Attributes.Duplicate();
                newAttributes.DeleteUserString(_originalObjectKey);
                if (newAttributes.SetUserString(entry.Key, entry.Value))
                {
                    doc.Objects.ModifyAttributes(rhinoObject.Id, newAttributes, false);
                    RhinoApp.WriteLine($"Updated user text for object {rhinoObject.Id}.");
                }
            }
            _originalObjectKey = null;
        }

        #endregion

        #region Section: Leader Options

        private Expander _viewportDisplayExpander;
        private CheckBox _displayIfcClassCheckBox;
        private CheckBox _displayMaterialCheckBox;
        private TextBox _ifcLeaderLengthTextBox;
        private TextBox _ifcLeaderAngleTextBox;
        private TextBox _materialLeaderLengthTextBox;
        private TextBox _materialLeaderAngleTextBox;

        private Control CreateDisplayOptionsLayout()
        {
            // Create controls
            _displayIfcClassCheckBox = new CheckBox { Text = "Display IFC Class" };
            _ifcLeaderLengthTextBox = new TextBox { Text = _ifcClassDisplayConduit.LeaderLength.ToString(), Width = 50 };
            _ifcLeaderAngleTextBox = new TextBox { Text = _ifcClassDisplayConduit.LeaderAngle.ToString(), Width = 50 };
            var applyIfcButton = new Button { Text = "Apply" };
            _displayMaterialCheckBox = new CheckBox { Text = "Display Material" };
            _materialLeaderLengthTextBox = new TextBox { Text = _materialDisplayConduit.LeaderLength.ToString(), Width = 50 };
            _materialLeaderAngleTextBox = new TextBox { Text = _materialDisplayConduit.LeaderAngle.ToString(), Width = 50 };
            var applyMaterialButton = new Button { Text = "Apply" };

            // Register events
            applyIfcButton.Click += OnApplyIfcDisplaySettingsClick;
            applyMaterialButton.Click += OnApplyMaterialDisplaySettingsClick;
            _displayIfcClassCheckBox.CheckedChanged += (s, e) => { PluginSettingsManager.SetBool(SettingKeys.DisplayIfcClass, _displayIfcClassCheckBox.Checked ?? false); OnDisplayAttributeTextCheckedChanged(s, e); };
            _displayMaterialCheckBox.CheckedChanged += (s, e) => { PluginSettingsManager.SetBool(SettingKeys.DisplayMaterial, _displayMaterialCheckBox.Checked ?? false); OnDisplayAttributeTextCheckedChanged(s, e); };

            // Assemble layout
            var layout = new DynamicLayout { Spacing = new Size(6, 6) };
            layout.AddRow(_displayIfcClassCheckBox);
            layout.AddRow(new Label { Text = "Leader Length Multi.:" }, _ifcLeaderLengthTextBox, new Label { Text = "Leader Angle:" }, _ifcLeaderAngleTextBox, null, applyIfcButton);
            layout.AddRow(_displayMaterialCheckBox);
            layout.AddRow(new Label { Text = "Leader Length Multi.:" }, _materialLeaderLengthTextBox, new Label { Text = "Leader Angle:" }, _materialLeaderAngleTextBox, null, applyMaterialButton);

            _viewportDisplayExpander = new Expander { Header = new Label { Text = "Leader Options", Style = "bold_label" }, Content = layout };
            _viewportDisplayExpander.ExpandedChanged += (s, e) => PluginSettingsManager.SetBool(SettingKeys.ExpanderLeaderOptions, _viewportDisplayExpander.Expanded);

            return _viewportDisplayExpander;
        }

        private void OnDisplayAttributeTextCheckedChanged(object sender, EventArgs e)
        {
            UpdateAllConduits();
            RhinoDoc.ActiveDoc?.Views.Redraw();
        }

        private void OnApplyIfcDisplaySettingsClick(object sender, EventArgs e)
        {
            if (ValidateAndApplyConduitSettings(_ifcLeaderLengthTextBox, _ifcLeaderAngleTextBox, _ifcClassDisplayConduit))
            {
                PluginSettingsManager.SetString(SettingKeys.IfcLeaderLength, _ifcLeaderLengthTextBox.Text);
                PluginSettingsManager.SetString(SettingKeys.IfcLeaderAngle, _ifcLeaderAngleTextBox.Text);
                RhinoDoc.ActiveDoc?.Views.Redraw();
            }
        }

        private void OnApplyMaterialDisplaySettingsClick(object sender, EventArgs e)
        {
            if (ValidateAndApplyConduitSettings(_materialLeaderLengthTextBox, _materialLeaderAngleTextBox, _materialDisplayConduit))
            {
                PluginSettingsManager.SetString(SettingKeys.MaterialLeaderLength, _materialLeaderLengthTextBox.Text);
                PluginSettingsManager.SetString(SettingKeys.MaterialLeaderAngle, _materialLeaderAngleTextBox.Text);
                RhinoDoc.ActiveDoc?.Views.Redraw();
            }
        }

        #endregion

        #region Section: Document User Text

        private Expander _documentUserTextExpander;
        private readonly GridView<DocumentUserTextEntry> _docUserTextGridView = new GridView<DocumentUserTextEntry> { ShowHeader = true, Height = 100, Width = 450 };

        private Control CreateDocumentUserTextLayout()
        {
            // Setup grid columns if they don't exist
            if (_docUserTextGridView.Columns.Count == 0)
            {
                _docUserTextGridView.Columns.Add(new GridColumn { HeaderText = "Key", DataCell = new TextBoxCell(nameof(DocumentUserTextEntry.Key)), Editable = false, Width = 150 });
                _docUserTextGridView.Columns.Add(new GridColumn { HeaderText = "Value", DataCell = new TextBoxCell(nameof(DocumentUserTextEntry.Value)), Editable = false, Width = 280 });
            }

            // Create controls
            var refreshButton = new Button { Text = "Refresh User Text" };

            // Register events
            refreshButton.Click += (s, e) => UpdatePanelData();

            // Assemble layout
            var layout = new DynamicLayout { Spacing = new Size(6, 6) };
            layout.AddRow(new Scrollable { Content = _docUserTextGridView, Width = 450, Height = 120 });
            layout.AddRow(refreshButton);

            _documentUserTextExpander = new Expander { Header = new Label { Text = "Document User Text", Style = "bold_label" }, Content = layout };
            _documentUserTextExpander.ExpandedChanged += (s, e) => PluginSettingsManager.SetBool(SettingKeys.ExpanderDocumentUserText, _documentUserTextExpander.Expanded);

            return _documentUserTextExpander;
        }

        #endregion

        #region Section: Advanced Settings

        private Expander _advancedSettingsExpander;
        private TextBox _materialCsvPathTextBox;
        private TextBox _ifcCsvPathTextBox;
        private TextBox _customCsvPathTextBox;

        private Control CreateAdvancedSettingsLayout()
        {
            // Create controls
            int textBoxWidth = 300;
            _ifcCsvPathTextBox = new TextBox { Width = textBoxWidth };
            _materialCsvPathTextBox = new TextBox { Width = textBoxWidth };
            _customCsvPathTextBox = new TextBox { Width = textBoxWidth };
            var saveButton = new Button
            {
                Text = "Save/Reload List",
                Image = BytesToEtoBitmap(Resources.btn_refreshList256, new Size(18, 18)),
                ImagePosition = ButtonImagePosition.Left,
                MinimumSize = Size.Empty
            };

            // Register events
            saveButton.Click += OnSaveAdvancedSettingsClick;

            // Assemble layout
            var layout = new DynamicLayout { Spacing = new Size(6, 6) };
            layout.AddRow(new Label { Text = "IFC Class CSV Path/URL:", VerticalAlignment = VerticalAlignment.Center }, _ifcCsvPathTextBox);
            layout.AddRow(new Label { Text = "Material CSV Path/URL:", VerticalAlignment = VerticalAlignment.Center }, _materialCsvPathTextBox);
            layout.AddRow(new Label { Text = "Custom Def. CSV Path/URL:", VerticalAlignment = VerticalAlignment.Center }, _customCsvPathTextBox);
            layout.AddRow(null, saveButton);

            _advancedSettingsExpander = new Expander { Header = new Label { Text = "Advanced Settings / Data Paths", Style = "bold_label" }, Content = layout };
            _advancedSettingsExpander.ExpandedChanged += (s, e) => PluginSettingsManager.SetBool(SettingKeys.ExpanderAdvancedSettings, _advancedSettingsExpander.Expanded);

            return _advancedSettingsExpander;
        }

        private void OnSaveAdvancedSettingsClick(object sender, EventArgs e)
        {
            PluginSettingsManager.SetString(SettingKeys.IfcClassCsvPath, _ifcCsvPathTextBox.Text);
            PluginSettingsManager.SetString(SettingKeys.MaterialCsvPath, _materialCsvPathTextBox.Text);
            PluginSettingsManager.SetString(SettingKeys.CustomCsvPath, _customCsvPathTextBox.Text);
            RhinoApp.WriteLine("Advanced settings saved.");
            OnRefreshListsClick(sender, e);
        }

        private void OnRefreshListsClick(object sender, EventArgs e)
        {
            RhinoApp.WriteLine("Refreshing IFC, Material, and Custom Attribute lists...");
            ReloadIfcClassList();
            ReloadMaterialList();
            ReloadCustomAttributeList();
            RhinoApp.WriteLine("Lists refreshed successfully.");
        }

        #endregion

        #region Global Event Handlers

        private void OnDocumentStateChanged(object sender, EventArgs e) => UpdatePanelDataSafe();
        private void OnDocumentChanged(object sender, EventArgs e)
        {
            _needsRefresh = true;
            ReloadIfcClassList();
            ReloadMaterialList();
            ReloadCustomAttributeList();
        }
        private void OnRhinoIdle(object sender, EventArgs e)
        {
            if (_needsRefresh)
            {
                _needsRefresh = false;
                UpdatePanelData();
            }
        }
        private void OnObjectAdded(object sender, RhinoObjectEventArgs e)
        {
            if (_isAssigningOnCreate) return;

            var doc = e.TheObject.Document;
            if (doc == null) return;

            bool assignDef = _assignDefOnCreateCheckBox?.Checked == true;
            bool assignCustom = _assignCustomOnCreateCheckBox?.Checked == true;
            if (!assignDef && !assignCustom) return;

            _isAssigningOnCreate = true;
            try
            {
                var newAttributes = e.TheObject.Attributes.Duplicate();
                bool attributesModified = false;

                if (assignDef)
                {
                    string primaryClass = (_ifcDropdown.SelectedValue as ListItem)?.Text;
                    if (!string.IsNullOrWhiteSpace(primaryClass))
                    {
                        string secondaryClass = (_ifcSubclassDropdown.SelectedValue as ListItem)?.Text;
                        string combinedClass = !string.IsNullOrWhiteSpace(secondaryClass) ? $"{primaryClass}, {secondaryClass}" : primaryClass;

                        if (newAttributes.SetUserString(IfcClassKey, combinedClass))
                            attributesModified = true;
                    }

                    string materialValue = (_materialDropdown.SelectedValue as ListItem)?.Text;
                    if (!string.IsNullOrWhiteSpace(materialValue))
                    {
                        if (newAttributes.SetUserString(PrimaryMaterialKey, materialValue))
                            attributesModified = true;
                    }
                }

                if (assignCustom)
                {
                    string key = (_customKeyDropdown.SelectedValue as ListItem)?.Text;
                    string value = (_customValueDropdown.SelectedValue as ListItem)?.Text;
                    if (!string.IsNullOrWhiteSpace(key) && !string.IsNullOrWhiteSpace(value))
                    {
                        var keysToRemove = newAttributes.GetUserStrings().AllKeys
                                             .Where(k => k.StartsWith(CustomAttributePrefix))
                                             .ToList();
                        foreach (var keyToRemove in keysToRemove)
                        {
                            newAttributes.DeleteUserString(keyToRemove);
                        }

                        string formattedKey = $"{CustomAttributePrefix}{key}";
                        if (newAttributes.SetUserString(formattedKey, value))
                            attributesModified = true;
                    }
                }

                if (attributesModified)
                {
                    doc.Objects.ModifyAttributes(e.ObjectId, newAttributes, false);
                }
            }
            finally
            {
                _isAssigningOnCreate = false;
            }
        }

        #endregion

        #region Data Loading & Panel Updates

        private void LoadSettings()
        {
            // Load checkbox states
            _showAllObjectsCheckBox.Checked = PluginSettingsManager.GetBool(SettingKeys.ShowAllObjects, true);
            _showUnassignedCheckBox.Checked = PluginSettingsManager.GetBool(SettingKeys.ShowUnassigned, true);
            _groupByMaterialCheckBox.Checked = PluginSettingsManager.GetBool(SettingKeys.GroupByMaterial, false);
            _groupByClassCheckBox.Checked = PluginSettingsManager.GetBool(SettingKeys.GroupByClass, true);
            _aggregateCheckBox.Checked = PluginSettingsManager.GetBool(SettingKeys.AggregateResults, false);
            _displayIfcClassCheckBox.Checked = PluginSettingsManager.GetBool(SettingKeys.DisplayIfcClass, false);
            _displayMaterialCheckBox.Checked = PluginSettingsManager.GetBool(SettingKeys.DisplayMaterial, false);
            _assignDefOnCreateCheckBox.Checked = PluginSettingsManager.GetBool(SettingKeys.AssignDefOnCreate, false);
            _assignCustomOnCreateCheckBox.Checked = PluginSettingsManager.GetBool(SettingKeys.AssignCustomOnCreate, false);

            // Load conduit text box values
            _ifcLeaderLengthTextBox.Text = _ifcClassDisplayConduit.LeaderLength.ToString();
            _ifcLeaderAngleTextBox.Text = _ifcClassDisplayConduit.LeaderAngle.ToString();
            _materialLeaderLengthTextBox.Text = _materialDisplayConduit.LeaderLength.ToString();
            _materialLeaderAngleTextBox.Text = _materialDisplayConduit.LeaderAngle.ToString();

            // Load CSV paths
            _ifcCsvPathTextBox.Text = PluginSettingsManager.GetString(SettingKeys.IfcClassCsvPath, string.Empty);
            _materialCsvPathTextBox.Text = PluginSettingsManager.GetString(SettingKeys.MaterialCsvPath, string.Empty);
            _customCsvPathTextBox.Text = PluginSettingsManager.GetString(SettingKeys.CustomCsvPath, string.Empty);

            // Load other UI states
            PluginSettingsManager.LoadColumnLayout(_userTextGridView.Columns);
            _utilityButtonsExpander.Expanded = PluginSettingsManager.GetBool(SettingKeys.ExpanderUtilityButtons, true);
            _definitionExpander.Expanded = PluginSettingsManager.GetBool(SettingKeys.ExpanderDefinition, true);
            _customDefinitionExpander.Expanded = PluginSettingsManager.GetBool(SettingKeys.ExpanderCustomDefinition, true);
            _attributeGridExpander.Expanded = PluginSettingsManager.GetBool(SettingKeys.ExpanderAttributeGrid, true);
            _objectAttributesExpander.Expanded = PluginSettingsManager.GetBool(SettingKeys.ExpanderObjectAttributes, true);
            _viewportDisplayExpander.Expanded = PluginSettingsManager.GetBool(SettingKeys.ExpanderLeaderOptions, true);
            _documentUserTextExpander.Expanded = PluginSettingsManager.GetBool(SettingKeys.ExpanderDocumentUserText, true);
            _advancedSettingsExpander.Expanded = PluginSettingsManager.GetBool(SettingKeys.ExpanderAdvancedSettings, true);
        }

        private void UpdatePanelDataSafe()
        {
            if (!_isSyncingSelection) UpdatePanelData();
        }
        private void UpdatePanelData() => Application.Instance.AsyncInvoke(UpdateUserTextGrid);
        private void UpdateUserTextGrid()
        {
            var doc = RhinoDoc.ActiveDoc;
            UpdateDocumentUserTextGrid(doc);
            UpdateObjectUserTextGrid(doc);

            if (doc == null || _materialLcaData == null)
            {
                _userTextGridView.DataStore = null;
                _totalLcaLabel.Text = "Total LCA: 0.00";
                return;
            }

            var initialObjects = (_showAllObjectsCheckBox.Checked == true
              ? doc.Objects.Where(obj => obj.IsSelectable(true, false, false, true) && obj.Attributes.Visible && doc.Layers[obj.Attributes.LayerIndex].IsVisible)
              : doc.Objects.GetSelectedObjects(false, false)).ToList();

            if (!initialObjects.Any())
            {
                _userTextGridView.DataStore = null;
                _totalLcaLabel.Text = "Total LCA: 0.00";
                return;
            }

            var processedEntries = new List<ProcessedObjectEntry>();
            foreach (var topLevelObject in initialObjects)
            {
                processedEntries.AddRange(FlattenObjects(topLevelObject, doc));
            }

            List<UserTextEntry> gridData;
            if (_aggregateCheckBox.Checked == true)
            {
                var aggregatedData = processedEntries
                  .GroupBy(p => new { Material = p.EffectiveAttributes.GetUserString(PrimaryMaterialKey) ?? "N/A", IfcClass = p.EffectiveAttributes.GetUserString(IfcClassKey) ?? "N/A", p.QuantityType, p.GroupName })
                  .Select(g =>
                  {
                      var totalEffectiveQuantity = g.Sum(p => p.Quantity * p.QuantityMultiplier);
                      _materialLcaData.TryGetValue(g.Key.Material, out MaterialData materialData);
                      var totalLca = 0.0;
                      if (materialData != null && g.Key.QuantityType != "N/A")
                      {
                          totalLca = totalEffectiveQuantity * materialData.Lca;
                      }

                      return new UserTextEntry
                      {
                          GroupName = g.Key.GroupName,
                          Value = g.Key.Material,
                          IfcClass = g.Key.IfcClass,
                          ReferenceQuantity = materialData?.ReferenceQuantity ?? 0,
                          ReferenceUnit = materialData?.ReferenceUnit ?? "",
                          ReferenceLca = materialData?.Lca ?? 0,
                          Count = g.Count(),
                          Quantity = totalEffectiveQuantity,
                          QuantityUnit = g.First().QuantityUnit,
                          QuantityType = g.Key.QuantityType,
                          Lca = totalLca,
                          ObjectIds = g.Select(p => p.TopLevelId).Distinct().ToList()
                      };
                  });
                gridData = SortAndFormatData(aggregatedData.ToList());
            }
            else
            {
                var data = processedEntries.Select(p =>
                {
                    var material = p.EffectiveAttributes.GetUserString(PrimaryMaterialKey) ?? "N/A";
                    double lca = 0.0;
                    _materialLcaData.TryGetValue(material, out MaterialData materialData);
                    if (materialData != null && p.QuantityType != "N/A") lca = (p.Quantity * p.QuantityMultiplier) * materialData.Lca;

                    return new UserTextEntry
                    {
                        GroupName = p.GroupName,
                        Value = material,
                        IfcClass = p.EffectiveAttributes.GetUserString(IfcClassKey) ?? "N/A",
                        ReferenceQuantity = materialData?.ReferenceQuantity ?? 0,
                        ReferenceUnit = materialData?.ReferenceUnit ?? "",
                        ReferenceLca = materialData?.Lca ?? 0,
                        Count = 1,
                        Quantity = p.Quantity,
                        QuantityUnit = p.QuantityUnit,
                        QuantityType = p.QuantityType,
                        QuantityMultiplier = p.QuantityMultiplier,
                        Lca = lca,
                        ObjectIds = new List<Guid> { p.TopLevelId }
                    };
                });
                gridData = SortAndFormatData(data.ToList());
            }

            if (_showUnassignedCheckBox.Checked == false)
            {
                gridData.RemoveAll(entry => entry.Value == "N/A");
            }

            _userTextGridView.DataStore = gridData;
            _totalLcaLabel.Text = $"Total LCA: {gridData.Sum(entry => entry.Lca).ToString("N2", CultureInfo.InvariantCulture)} kgCO2 eq";

            SyncGridSelectionWithRhino();
            UpdateAllConduits();
            doc.Views.Redraw();
        }
        private void UpdateAllConduits()
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;

            _ifcClassDisplayConduit.DataToDraw.Clear();
            _materialDisplayConduit.DataToDraw.Clear();

            bool ifcEnabled = _displayIfcClassCheckBox.Checked ?? false;
            bool materialEnabled = _displayMaterialCheckBox.Checked ?? false;

            if (!ifcEnabled && !materialEnabled)
            {
                _ifcClassDisplayConduit.Enabled = false;
                _materialDisplayConduit.Enabled = false;
                return;
            }

            var allObjects = doc.Objects.Where(obj => obj.IsSelectable(true, false, false, true) && obj.Attributes.Visible && doc.Layers[obj.Attributes.LayerIndex].IsVisible);

            foreach (var rhinoObject in allObjects)
            {
                if (ifcEnabled)
                {
                    var ifcClass = rhinoObject.Attributes.GetUserString(IfcClassKey);
                    if (!string.IsNullOrWhiteSpace(ifcClass))
                    {
                        _ifcClassDisplayConduit.DataToDraw.Add(new ConduitDrawData { ObjectId = rhinoObject.Id, Text = ifcClass });
                    }
                }

                if (materialEnabled)
                {
                    var material = rhinoObject.Attributes.GetUserString(PrimaryMaterialKey);
                    if (!string.IsNullOrWhiteSpace(material))
                    {
                        _materialDisplayConduit.DataToDraw.Add(new ConduitDrawData { ObjectId = rhinoObject.Id, Text = material });
                    }
                }
            }

            _ifcClassDisplayConduit.Enabled = ifcEnabled;
            _materialDisplayConduit.Enabled = materialEnabled;
        }
        private void ReloadIfcClassList()
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;
            _ifcClasses = CsvReader.ReadIfcClassesDynamic(doc);
            PopulatePrimaryIfcDropdown();
        }
        private void ReloadMaterialList()
        {
            var doc = RhinoDoc.ActiveDoc;
            var materialList = CsvReader.ReadMaterialLcaDataDynamic(doc);
            _materialLcaData = materialList.GroupBy(m => m.MaterialName).ToDictionary(g => g.Key, g => g.First());
            PopulateMaterialDropdown();
            UpdatePanelData();
        }

        private void ReloadCustomAttributeList()
        {
            var doc = RhinoDoc.ActiveDoc;
            CsvReader.ReadCustomAttributeListsDynamic(doc, out _customAttributeKeys, out _customAttributeValues);
            PopulateCustomKeyDropdown();
            PopulateCustomValueDropdown();
        }

        private void PopulateCustomKeyDropdown()
        {
            var selectedText = (_customKeyDropdown.SelectedValue as ListItem)?.Text;
            _allCustomKeyItems.Clear();
            if (_customAttributeKeys != null)
            {
                foreach (var key in _customAttributeKeys)
                    _allCustomKeyItems.Add(new ListItem { Text = key });
            }
            _filteredCustomKeyItems.Filter = null;
            _filteredCustomKeyItems.Refresh();
            _customKeyDropdown.SelectedValue = _allCustomKeyItems.FirstOrDefault(i => i.Text == selectedText);
        }
        private void PopulateCustomValueDropdown()
        {
            var selectedText = (_customValueDropdown.SelectedValue as ListItem)?.Text;
            _allCustomValueItems.Clear();
            if (_customAttributeValues != null)
            {
                foreach (var value in _customAttributeValues)
                    _allCustomValueItems.Add(new ListItem { Text = value });
            }
            _filteredCustomValueItems.Filter = null;
            _filteredCustomValueItems.Refresh();
            _customValueDropdown.SelectedValue = _allCustomValueItems.FirstOrDefault(i => i.Text == selectedText);
        }
        private void PopulateMaterialDropdown()
        {
            var selectedText = (_materialDropdown.SelectedValue as ListItem)?.Text;
            _allMaterialListItems.Clear();
            if (_materialLcaData != null)
            {
                foreach (var materialName in _materialLcaData.Keys)
                    _allMaterialListItems.Add(new ListItem { Text = materialName });
            }
            _filteredMaterialItems.Filter = null;
            _filteredMaterialItems.Refresh();
            _materialDropdown.SelectedValue = _allMaterialListItems.FirstOrDefault(i => i.Text == selectedText);
        }
        private void PopulatePrimaryIfcDropdown()
        {
            var selectedText = (_ifcDropdown.SelectedValue as ListItem)?.Text;
            _allIfcClassItems.Clear();
            foreach (var primaryClass in _ifcClasses.Keys)
            {
                _allIfcClassItems.Add(new ListItem { Text = primaryClass });
            }
            _filteredIfcClassItems.Filter = null;
            _filteredIfcClassItems.Refresh();
            _ifcDropdown.SelectedValue = _allIfcClassItems.FirstOrDefault(i => i.Text == selectedText);
            UpdateSubclassDropdown();
        }

        private void UpdateSubclassDropdown()
        {
            var selectedText = (_ifcSubclassDropdown.SelectedValue as ListItem)?.Text;
            _allIfcSubclassItems.Clear();

            if (_ifcDropdown.SelectedValue is ListItem selectedItem && !string.IsNullOrEmpty(selectedItem.Text))
            {
                string primaryClass = selectedItem.Text;
                if (_ifcClasses.TryGetValue(primaryClass, out List<string> subclasses) && subclasses.Any())
                {
                    foreach (var subclass in subclasses)
                        _allIfcSubclassItems.Add(new ListItem { Text = subclass });
                    _ifcSubclassDropdown.Enabled = true;
                }
                else
                {
                    _ifcSubclassDropdown.Enabled = false;
                }
            }
            else
            {
                var allSubclasses = _ifcClasses.Values
                    .SelectMany(sublist => sublist)
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Distinct();

                foreach (var subclass in allSubclasses)
                    _allIfcSubclassItems.Add(new ListItem { Text = subclass });

                _ifcSubclassDropdown.Enabled = true;
            }

            _filteredIfcSubclassItems.Filter = null;
            _filteredIfcSubclassItems.Refresh();
            _ifcSubclassDropdown.SelectedValue = _allIfcSubclassItems.FirstOrDefault(i => i.Text == selectedText);
        }

        private void UpdateDocumentUserTextGrid(RhinoDoc doc)
        {
            if (doc == null || _docUserTextGridView == null)
            {
                if (_docUserTextGridView != null) _docUserTextGridView.DataStore = null;
                return;
            }
            var docStrings = doc.Strings;
            if (docStrings.Count == 0)
            {
                _docUserTextGridView.DataStore = null;
            }
            else
            {
                var data = new List<DocumentUserTextEntry>(docStrings.Count);
                for (int i = 0; i < docStrings.Count; i++)
                {
                    var key = docStrings.GetKey(i);
                    if (!string.IsNullOrEmpty(key)) data.Add(new DocumentUserTextEntry { Key = key, Value = docStrings.GetValue(key) });
                }
                _docUserTextGridView.DataStore = data.OrderBy(d => d.Key).ToList();
            }
        }

        private void UpdateObjectUserTextGrid(RhinoDoc doc)
        {
            if (doc == null)
            {
                _objectUserTextGridView.DataStore = null;
                return;
            }

            var selectedObjects = doc.Objects.GetSelectedObjects(false, false).ToList();

            var effectiveAttributeList = new List<ObjectAttributes>();
            if (selectedObjects.Any())
            {
                foreach (var obj in selectedObjects)
                {
                    effectiveAttributeList.AddRange(GetEffectiveAttributesFromObject(obj));
                }
            }

            // Allow editing only if a single, non-block object is selected.
            bool singleSelection = (selectedObjects.Count == 1 && !(selectedObjects[0] is InstanceObject));
            _objectKeyColumn.Editable = singleSelection;
            _objectValueColumn.Editable = singleSelection;

            int selectionCount = selectedObjects.Count;
            string selectionText;
            if (selectionCount == 0)
            {
                selectionText = "No Object Selected";
            }
            else if (selectionCount == 1)
            {
                selectionText = "1 Object Selected";
            }
            else
            {
                selectionText = $"{selectionCount} Objects Selected";
            }
            _selectionCountLabel.Text = selectionText;

            if (!effectiveAttributeList.Any())
            {
                _objectUserTextGridView.DataStore = null;
                return;
            }

            var allAttributes = new Dictionary<string, HashSet<string>>();

            foreach (var attributes in effectiveAttributeList)
            {
                var userStrings = attributes.GetUserStrings();
                foreach (var key in userStrings.AllKeys)
                {
                    if (!allAttributes.ContainsKey(key))
                    {
                        allAttributes[key] = new HashSet<string>();
                    }
                    allAttributes[key].Add(userStrings.Get(key) ?? string.Empty);
                }
            }

            var data = new List<ObjectUserTextEntry>();
            foreach (var kvp in allAttributes.OrderBy(item => item.Key))
            {
                string displayValue;
                var distinctValues = kvp.Value.ToList();

                if (distinctValues.Count == 1)
                {
                    displayValue = distinctValues[0];
                }
                else
                {
                    displayValue = string.Join(", ", distinctValues.Select(v => $"\"{v}\""));
                }
                data.Add(new ObjectUserTextEntry { Key = kvp.Key, Value = displayValue });
            }

            for (int i = 0; i < data.Count; i++)
            {
                data[i].SerialNumber = $"{i + 1}.";
            }

            _objectUserTextGridView.DataStore = data;
        }

        #endregion

        #region Core Logic & Helpers

        /// <summary>
        /// A helper class to hold the processed data for a single geometric object, including those inside blocks.
        /// </summary>
        private class ProcessedObjectEntry
        {
            public Guid TopLevelId { get; set; }
            public GeometryBase Geometry { get; set; }
            public ObjectAttributes EffectiveAttributes { get; set; }
            public double Quantity { get; set; }
            public string QuantityUnit { get; set; }
            public string QuantityType { get; set; }
            public double QuantityMultiplier { get; set; } = 1.0;
            public string GroupName { get; set; }
        }

        /// <summary>
        /// Recursively flattens a RhinoObject, handling nested blocks and transforming geometry.
        /// It returns a list of entries, each with final geometry, effective attributes, and top-level ID.
        /// </summary>
        private List<ProcessedObjectEntry> FlattenObjects(RhinoObject topLevelObject, RhinoDoc doc)
        {
            var results = new List<ProcessedObjectEntry>();
            // Determine the group or block name at the top level
            string name = "-";
            var groupIndices = topLevelObject.GetGroupList();
            if (groupIndices != null && groupIndices.Length > 0)
            {
                var group = doc.Groups.FindIndex(groupIndices[0]);
                if (group != null && !string.IsNullOrWhiteSpace(group.Name))
                {
                    name = group.Name;
                }
            }
            else if (topLevelObject is InstanceObject io)
            {
                name = io.InstanceDefinition.Name;
            }
            RecursiveFlatten(topLevelObject, Transform.Identity, topLevelObject.Attributes, topLevelObject.Id, name, results, doc);
            return results;
        }

        private void RecursiveFlatten(RhinoObject currentObject, Transform parentTransform, ObjectAttributes parentAttributes, Guid topLevelId, string groupName, List<ProcessedObjectEntry> results, RhinoDoc doc)
        {
            // Combine the current transform with the parent's
            var currentTransform = parentTransform;

            // Start with a copy of the parent's attributes
            var effectiveAttributes = parentAttributes.Duplicate();

            // Instance-level attributes override parent attributes
            var currentUserStrings = currentObject.Attributes.GetUserStrings();
            foreach (var key in currentUserStrings.AllKeys)
            {
                effectiveAttributes.SetUserString(key, currentUserStrings.Get(key));
            }

            if (currentObject is InstanceObject io)
            {
                currentTransform = parentTransform * io.InstanceXform;
                foreach (var subObject in io.InstanceDefinition.GetObjects())
                {
                    // Recurse into the block definition
                    RecursiveFlatten(subObject, currentTransform, effectiveAttributes, topLevelId, groupName, results, doc);
                }
            }
            else // It's a geometric object
            {
                var transformedGeometry = currentObject.Geometry.Duplicate();
                transformedGeometry.Transform(parentTransform);

                TryComputeQuantity(transformedGeometry, doc, out double qty, out string unit, out string qtyType);

                var multiplierString = effectiveAttributes.GetUserString(QuantityMultiplierKey);
                double multiplier = 1.0;
                if (!string.IsNullOrEmpty(multiplierString) && double.TryParse(multiplierString, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedMultiplier) && parsedMultiplier > 0)
                {
                    multiplier = parsedMultiplier;
                }

                results.Add(new ProcessedObjectEntry
                {
                    TopLevelId = topLevelId,
                    Geometry = transformedGeometry,
                    EffectiveAttributes = effectiveAttributes,
                    Quantity = qty,
                    QuantityUnit = unit,
                    QuantityType = qtyType,
                    QuantityMultiplier = multiplier,
                    GroupName = groupName
                });
            }
        }

        /// <summary>
        /// Recursively gets the effective attributes of an object and all its children if it's a block.
        /// </summary>
        private List<ObjectAttributes> GetEffectiveAttributesFromObject(RhinoObject obj)
        {
            var results = new List<ObjectAttributes>();
            RecursiveGetEffectiveAttributes(obj, obj.Attributes, results);
            return results;
        }

        private void RecursiveGetEffectiveAttributes(RhinoObject currentObject, ObjectAttributes parentAttributes, List<ObjectAttributes> results)
        {
            var effectiveAttributes = parentAttributes.Duplicate();
            var currentUserStrings = currentObject.Attributes.GetUserStrings();
            foreach (var key in currentUserStrings.AllKeys)
            {
                effectiveAttributes.SetUserString(key, currentUserStrings.Get(key));
            }

            if (currentObject is InstanceObject io)
            {
                foreach (var subObject in io.InstanceDefinition.GetObjects())
                {
                    RecursiveGetEffectiveAttributes(subObject, effectiveAttributes, results);
                }
            }
            else
            {
                results.Add(effectiveAttributes);
            }
        }

        private bool ValidateAndApplyConduitSettings(TextBox lengthBox, TextBox angleBox, AttributeDisplayConduit conduit)
        {
            bool lengthValid = int.TryParse(lengthBox.Text, out int length) && length > 0 && length <= 50;
            if (!lengthValid)
            {
                RhinoApp.WriteLine("Invalid Leader Length. Please enter an integer between 1 and 50.");
                lengthBox.Text = conduit.LeaderLength.ToString();
                return false;
            }

            bool angleValid = int.TryParse(angleBox.Text, out int angle) && angle >= 0 && angle <= 360;
            if (!angleValid)
            {
                RhinoApp.WriteLine("Invalid Leader Angle. Please enter an integer between 0 and 360.");
                angleBox.Text = conduit.LeaderAngle.ToString();
                return false;
            }

            conduit.LeaderLength = length;
            conduit.LeaderAngle = angle;
            return true;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                PluginSettingsManager.SaveColumnLayout(_userTextGridView.Columns);
                _ifcClassDisplayConduit.Enabled = false;
                _materialDisplayConduit.Enabled = false;
            }
            base.Dispose(disposing);
        }
        private List<UserTextEntry> SortAndFormatData(List<UserTextEntry> data)
        {
            IEnumerable<UserTextEntry> sortedData;
            if (_groupByClassCheckBox.Checked == true) sortedData = data.OrderBy(d => d.IfcClass == "N/A").ThenBy(d => d.IfcClass).ThenBy(d => d.Value);
            else if (_groupByMaterialCheckBox.Checked == true) sortedData = data.OrderBy(d => d.Value == "N/A").ThenBy(d => d.Value).ThenBy(d => d.IfcClass);
            else sortedData = data.OrderBy(d => d.Value == "N/A").ThenBy(d => d.Value);

            var formattedList = new List<UserTextEntry>();
            string lastPrimaryGroup = null;
            int serialNumber = 1;
            foreach (var item in sortedData)
            {
                string currentPrimaryGroup = GetCurrentGroupingKey(item);
                bool isFirstInGroup = currentPrimaryGroup != lastPrimaryGroup;
                item.SerialNumber = isFirstInGroup ? $"{serialNumber++}." : "";
                if (!isFirstInGroup)
                {
                    if (_groupByClassCheckBox.Checked == true) item.IfcClass = "";
                    else if (_groupByMaterialCheckBox.Checked == true) item.Value = "";
                }
                formattedList.Add(item);
                lastPrimaryGroup = currentPrimaryGroup;
            }
            return formattedList;
        }
        private string GetCurrentGroupingKey(UserTextEntry item)
        {
            if (_groupByClassCheckBox.Checked == true) return item.IfcClass;
            if (_groupByMaterialCheckBox.Checked == true) return item.Value;
            return Guid.NewGuid().ToString();
        }
        private void SyncGridSelectionWithRhino()
        {
            if (_isSyncingSelection) return;
            _isSyncingSelection = true;
            var doc = RhinoDoc.ActiveDoc;
            if (doc != null)
            {
                var selectedObjectIds = doc.Objects.GetSelectedObjects(false, false).Select(o => o.Id).ToHashSet();
                var rowsToSelect = new List<int>();
                if (_userTextGridView.DataStore is List<UserTextEntry> dataStoreList)
                {
                    for (int i = 0; i < dataStoreList.Count; i++)
                    {
                        if (dataStoreList[i].ObjectIds.Any(id => selectedObjectIds.Contains(id))) rowsToSelect.Add(i);
                    }
                }
                _userTextGridView.SelectedRows = rowsToSelect;
            }
            _isSyncingSelection = false;
        }
        private static void AssignUserString(string key, string friendlyName, string value)
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;

            var selectedObjects = doc.Objects.GetSelectedObjects(false, false).ToList();
            if (!selectedObjects.Any())
            {
                RhinoApp.WriteLine("Please select one or more objects in the viewport.");
                return;
            }
            int updatedCount = 0;
            foreach (var rhinoObject in selectedObjects)
            {
                var newAttributes = rhinoObject.Attributes.Duplicate();
                if (newAttributes.SetUserString(key, value) && doc.Objects.ModifyAttributes(rhinoObject.Id, newAttributes, false))
                {
                    updatedCount++;
                }
            }
            if (updatedCount > 0)
            {
                RhinoApp.WriteLine($"Successfully assigned '{value}' to {updatedCount} object(s).");
                doc.Views.Redraw();
            }
        }
        private static void RemoveUserString(string key, string friendlyName)
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;
            var selectedObjects = doc.Objects.GetSelectedObjects(false, false).ToList();
            if (!selectedObjects.Any())
            {
                RhinoApp.WriteLine($"Please select one or more objects to remove the '{friendlyName}' definition from.");
                return;
            }
            int updatedCount = 0;
            foreach (var rhinoObject in selectedObjects)
            {
                if (rhinoObject?.Attributes.GetUserString(key) != null)
                {
                    var newAttributes = rhinoObject.Attributes.Duplicate();
                    newAttributes.DeleteUserString(key);
                    if (doc.Objects.ModifyAttributes(rhinoObject.Id, newAttributes, false)) updatedCount++;
                }
            }
            if (updatedCount > 0)
            {
                RhinoApp.WriteLine($"Successfully removed '{friendlyName}' definition from {updatedCount} object(s).");
                doc.Views.Redraw();
            }
            else
            {
                RhinoApp.WriteLine($"Selected objects do not have a '{friendlyName}' definition to remove.");
            }
        }
        private void ReorderGridColumns(bool groupByMaterial)
        {
            var ifcClassColumn = _userTextGridView.Columns.FirstOrDefault(c => c.HeaderText == "IfcClass");
            var materialNameColumn = _userTextGridView.Columns.FirstOrDefault(c => c.HeaderText == "Material (Name)");
            if (ifcClassColumn == null || materialNameColumn == null) return;
            int ifcIndex = _userTextGridView.Columns.IndexOf(ifcClassColumn);
            int materialIndex = _userTextGridView.Columns.IndexOf(materialNameColumn);
            if ((groupByMaterial && ifcIndex < materialIndex) || (!groupByMaterial && materialIndex < ifcIndex))
            {
                _userTextGridView.Columns.Move(ifcIndex, materialIndex);
            }
        }
        private static bool TryComputeQuantity(GeometryBase geo, RhinoDoc doc, out double quantity, out string unit, out string quantityType)
        {
            quantity = 0.0;
            unit = string.Empty;
            quantityType = "N/A";
            if (geo == null || doc == null) return false;
            var unitAbbreviation = doc.GetUnitSystemName(true, true, false, true);

            if (geo is Rhino.Geometry.Point)
            {
                quantityType = "Each";
                quantity = 1.0;
                unit = "ea.";
                return true;
            }
            if (TryGetVolume(geo, out quantity))
            {
                quantityType = "Volume";
                unit = $"{unitAbbreviation}³";
                return true;
            }
            if (TryGetArea(geo, out quantity))
            {
                quantityType = "Area";
                unit = $"{unitAbbreviation}²";
                return true;
            }
            if (TryGetLength(geo, out quantity))
            {
                quantityType = "Length";
                unit = unitAbbreviation;
                return true;
            }
            return false;
        }
        private static bool TryGetVolume(GeometryBase geo, out double quantity)
        {
            quantity = 0.0;
            switch (geo)
            {
                case Brep brep when brep.IsSolid: quantity = brep.GetVolume(); break;
                case Extrusion extrusion: var brepFromExtrusion = extrusion.ToBrep(); if (brepFromExtrusion != null && brepFromExtrusion.IsSolid) quantity = brepFromExtrusion.GetVolume(); break;
                case Mesh mesh when mesh.IsClosed: quantity = mesh.Volume(); break;
            }
            return quantity > 1e-9;
        }
        private static bool TryGetArea(GeometryBase geo, out double quantity)
        {
            quantity = 0.0;
            AreaMassProperties amp = null;
            switch (geo)
            {
                case Brep brep: amp = AreaMassProperties.Compute(brep); break;
                case Surface srf: amp = AreaMassProperties.Compute(srf); break;
                case Mesh mesh: amp = AreaMassProperties.Compute(mesh); break;
                case Curve curve when curve.IsPlanar(): amp = AreaMassProperties.Compute(curve); break;
            }
            if (amp != null) quantity = amp.Area;
            return quantity > 1e-9;
        }
        private static bool TryGetLength(GeometryBase geo, out double quantity)
        {
            quantity = 0.0;
            if (geo is Curve curve) quantity = curve.GetLength();
            return quantity > 1e-9;
        }
        private static Bitmap BytesToEtoBitmap(byte[] bytes, Size? desiredSize = null)
        {
            try
            {
                if (bytes == null || bytes.Length == 0) return null;
                using (var ms = new MemoryStream(bytes))
                {
                    var originalBitmap = new Bitmap(ms);
                    if (desiredSize.HasValue)
                    {
                        return new Bitmap(originalBitmap, desiredSize.Value.Width, desiredSize.Value.Height, ImageInterpolation.High);
                    }
                    return originalBitmap;
                }
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error creating bitmap from resource bytes: {ex.Message}");
                return null;
            }
        }

        #endregion

        #region Inner Classes

        private class UserTextEntry
        {
            public string SerialNumber { get; set; }
            public string GroupName { get; set; }
            public string IfcClass { get; set; }
            public string Value { get; set; }
            public double ReferenceQuantity { get; set; }
            public string ReferenceUnit { get; set; }
            public double ReferenceLca { get; set; }
            public string DisplayReference => !string.IsNullOrWhiteSpace(ReferenceUnit) ? $"{ReferenceQuantity:0.00} {ReferenceUnit}" : "N/A";
            public int Count { get; set; }
            public double Quantity { get; set; }
            public string QuantityUnit { get; set; }
            public string DisplayQuantity => !string.IsNullOrWhiteSpace(QuantityUnit) ? $"{Quantity:0.00} {QuantityUnit}" : "N/A";
            public string QuantityType { get; set; }
            public double QuantityMultiplier { get; set; } = 1.0;
            public double QuantityTotal => Quantity * QuantityMultiplier;
            public double Lca { get; set; }
            public List<Guid> ObjectIds { get; set; } = new List<Guid>();
        }
        private class DocumentUserTextEntry
        {
            public string Key { get; set; }
            public string Value { get; set; }
        }
        private class ObjectUserTextEntry
        {
            public string SerialNumber { get; set; }
            public string Key { get; set; }
            public string Value { get; set; }
        }
        private class ConduitDrawData
        {
            public Guid ObjectId { get; set; }
            public string Text { get; set; }
        }
        private class AttributeDisplayConduit : DisplayConduit
        {
            public List<ConduitDrawData> DataToDraw { get; set; } = new List<ConduitDrawData>();
            public int LeaderLength { get; set; } = 5;
            public int LeaderAngle { get; set; } = 45;

            protected override void PostDrawObjects(DrawEventArgs e)
            {
                if (!e.Display.Viewport.IsPerspectiveProjection) return;

                e.Display.PushDepthTesting(false);
                e.Display.PushDepthWriting(false);

                var textColor = System.Drawing.Color.Black;
                var backgroundColor = System.Drawing.Color.White;
                const string fontFace = "Arial";
                const int textPixelHeight = 10;

                foreach (var data in DataToDraw)
                {
                    var rhinoObject = e.RhinoDoc.Objects.FindId(data.ObjectId);
                    if (rhinoObject == null) continue;

                    var bbox = rhinoObject.Geometry.GetBoundingBox(true);
                    if (!bbox.IsValid) continue;

                    var location = bbox.Center;
                    var text = data.Text;

                    var worldToScreen = e.Display.Viewport.GetTransform(CoordinateSystem.World, CoordinateSystem.Screen);
                    var locationOnScreen = location;
                    locationOnScreen.Transform(worldToScreen);
                    var pointAboveOnScreen = new Point3d(locationOnScreen.X, locationOnScreen.Y - textPixelHeight, locationOnScreen.Z);

                    var screenToWorld = e.Display.Viewport.GetTransform(CoordinateSystem.Screen, CoordinateSystem.World);
                    pointAboveOnScreen.Transform(screenToWorld);
                    var worldHeight = location.DistanceTo(pointAboveOnScreen);

                    var camDir = e.Display.Viewport.CameraDirection;
                    var camUp = e.Display.Viewport.CameraUp;
                    var screenXAxis = Vector3d.CrossProduct(camDir, camUp);

                    double angleRad = LeaderAngle * Math.PI / 180.0;
                    var leaderDirection = (screenXAxis * Math.Cos(angleRad)) + (camUp * Math.Sin(angleRad));
                    leaderDirection.Unitize();

                    var leaderLength = worldHeight * LeaderLength;
                    var textLocation = location + (leaderDirection * leaderLength);

                    e.Display.DrawArrow(new Line(textLocation, location), textColor);

                    var tempTextEntity = new Text3d(text, Plane.WorldXY, worldHeight) { FontFace = fontFace };
                    var tempBbox = tempTextEntity.BoundingBox;
                    var worldWidth = tempBbox.Max.X - tempBbox.Min.X;

                    var padding = worldHeight * 0.2;
                    var backgroundWidth = worldWidth + (2 * padding);
                    var backgroundHeight = worldHeight + (2 * padding);
                    var backgroundOrigin = textLocation - (screenXAxis * padding) - (camUp * padding);

                    var corner1 = backgroundOrigin;
                    var corner2 = backgroundOrigin + (screenXAxis * backgroundWidth);
                    var corner3 = backgroundOrigin + (screenXAxis * backgroundWidth) + (camUp * backgroundHeight);
                    var corner4 = backgroundOrigin + (camUp * backgroundHeight);

                    var backgroundCorners = new Point3d[] { corner1, corner2, corner3, corner4 };
                    e.Display.DrawPolygon(backgroundCorners, backgroundColor, true);

                    var textPlane = new Plane(textLocation, screenXAxis, camUp);
                    e.Display.Draw3dText(text, textColor, textPlane, worldHeight, fontFace);
                }

                e.Display.PopDepthTesting();
                e.Display.PopDepthWriting();
            }
        }

        #endregion
    }
}