// MyProjectPanel.cs

using Eto.Drawing;
using Eto.Forms;
using MyProject.Properties;
using Rhino;
using Rhino.Display;
using Rhino.DocObjects;
using Rhino.Geometry;
using Rhino.UI.Controls;
using System;
using System.Collections.Generic;
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
        #region Constants
        private const string MaterialKey = "STR_MATERIAL";
        private const string IfcClassKey = "STR_IFC_CLASS";
        private const string QuantityMultiplierKey = "STR_QTY_MULTIPLIER";
        private const string CustomAttributePrefix = "CUSTOM_";
        #endregion

        #region Private Fields

        // Data Storage
        private Dictionary<string, List<string>> _ifcClasses = new Dictionary<string, List<string>>();
        private Dictionary<string, MaterialData> _materialLcaData;
        private List<string> _customAttributeKeys;
        private List<string> _customAttributeValues;

        // UI Controls
        private readonly GridView<UserTextEntry> _userTextGridView = new GridView<UserTextEntry> { ShowHeader = true, AllowMultipleSelection = true, Width = 450 };
        private readonly GridView<DocumentUserTextEntry> _docUserTextGridView = new GridView<DocumentUserTextEntry> { ShowHeader = true, Height = 100, Width = 450 };
        private readonly GridView<ObjectUserTextEntry> _objectUserTextGridView = new GridView<ObjectUserTextEntry> { ShowHeader = true, Height = 150, Width = 450 };
        private readonly GridColumn _qtyMultiplierColumn = new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.QuantityMultiplier)) { TextAlignment = TextAlignment.Right }, HeaderText = "Qty. Multiplier", Editable = true, Width = 90 };
        private readonly Label _totalLcaLabel = new Label { Text = "Total LCA: 0.00" };
        private Label _selectedObjectAttributesHeaderLabel;
        private readonly ComboBox _ifcDropdown = new ComboBox { Width = 90, AutoComplete = true };
        private readonly ComboBox _ifcSubclassDropdown = new ComboBox { Width = 90, AutoComplete = true };
        private readonly ComboBox _materialDropdown = new ComboBox { Width = 90, AutoComplete = true };
        private ComboBox _customKeyDropdown;
        private ComboBox _customValueDropdown;
        private readonly CheckBox _showAllObjectsCheckBox = new CheckBox { Text = "Show All/Selected" };
        private readonly CheckBox _showUnassignedCheckBox = new CheckBox { Text = "Show/Hide Unassigned" };
        private readonly CheckBox _groupByMaterialCheckBox = new CheckBox { Text = "Group by Material" };
        private readonly CheckBox _groupByClassCheckBox = new CheckBox { Text = "Group by Class" };
        private readonly CheckBox _aggregateCheckBox = new CheckBox { Text = "Aggregate" };
        private CheckBox _displayIfcClassCheckBox;
        private CheckBox _displayMaterialCheckBox;
        private CheckBox _assignDefOnCreateCheckBox;
        private CheckBox _assignCustomOnCreateCheckBox;

        private TextBox _ifcLeaderLengthTextBox;
        private TextBox _ifcLeaderAngleTextBox;
        private TextBox _materialLeaderLengthTextBox;
        private TextBox _materialLeaderAngleTextBox;

        // Conduits
        private readonly AttributeDisplayConduit _ifcClassDisplayConduit;
        private readonly AttributeDisplayConduit _materialDisplayConduit;

        // State Flags
        private bool _isAssigningOnCreate = false;
        private bool _isSyncingSelection = false;
        private bool _needsRefresh = false;
        private bool _isMouseOverGrid = false;

        #endregion

        #region Constructor

        public MyProjectPanel()
        {
            // --- MODIFICATION START: Initialize conduits with saved or default values ---
            _ifcClassDisplayConduit = new AttributeDisplayConduit
            {
                LeaderLength = int.Parse(PluginSettingsManager.GetString(SettingKeys.IfcLeaderLength, "5")),
                LeaderAngle = int.Parse(PluginSettingsManager.GetString(SettingKeys.IfcLeaderAngle, "45"))
            };
            _materialDisplayConduit = new AttributeDisplayConduit
            {
                LeaderLength = int.Parse(PluginSettingsManager.GetString(SettingKeys.MaterialLeaderLength, "8")),
                LeaderAngle = int.Parse(PluginSettingsManager.GetString(SettingKeys.MaterialLeaderAngle, "45"))
            };
            // --- MODIFICATION END ---

            InitializeLayout();
            RegisterEventHandlers();

            // --- MODIFICATION START: Load settings after UI is initialized ---
            LoadSettings();
            // --- MODIFICATION END ---

            ReloadIfcClassList();
            ReloadMaterialList();
            ReloadCustomAttributeList();
            UpdatePanelData();
        }
        #endregion

        #region UI Initialization


        private void InitializeLayout()
        {
            Styles.Add<Label>("bold_label", label =>
            {
                label.Font = SystemFonts.Bold();
                _totalLcaLabel.Font = label.Font;
            });

            var mainLayout = new DynamicLayout { Spacing = new Size(6, 6), Padding = new Padding(10) };

            _selectedObjectAttributesHeaderLabel = new Label { Text = "Selected Object(s) Attributes", Style = "bold_label" };

            mainLayout.Add(new Expander { Header = new Label { Text = "Utility Buttons", Style = "bold_label" }, Content = CreateUtilityButtonsLayout(), Expanded = true });
            mainLayout.Add(new Expander { Header = new Label { Text = "Definition", Style = "bold_label" }, Content = CreateDefinitionLayout(), Expanded = true });
            mainLayout.Add(new Expander { Header = new Label { Text = "Custom Definition", Style = "bold_label" }, Content = CreateCustomDefinitionLayout(), Expanded = true });
            mainLayout.Add(new Expander { Header = new Label { Text = "Attribute User Text", Style = "bold_label" }, Content = CreateAttributeGridLayout(), Expanded = true });
            mainLayout.Add(new Expander { Header = _selectedObjectAttributesHeaderLabel, Content = CreateObjectUserTextLayout(), Expanded = true });
            mainLayout.Add(new Expander { Header = new Label { Text = "Viewport Display", Style = "bold_label" }, Content = CreateDisplayOptionsLayout(), Expanded = true });
            mainLayout.Add(new Expander { Header = new Label { Text = "Document User Text", Style = "bold_label" }, Content = CreateDocumentUserTextLayout(), Expanded = false });
            mainLayout.Add(null, true);

            Content = new Scrollable { Content = mainLayout, Border = BorderType.None };
            MinimumSize = new Size(400, 450);
        }


        private Control CreateUtilityButtonsLayout()
        {
            var structureIcon = BytesToEtoBitmap(Resources.btn_strLogo256, new Size(18, 18));
            var btnStructure = new Button { Image = structureIcon, ToolTip = "Visit str-ucture.com", MinimumSize = Size.Empty, ImagePosition = ButtonImagePosition.Left };
            btnStructure.Click += (s, e) =>
            {
                var url = "http://www.str-ucture.com";
                try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
                catch (Exception ex) { RhinoApp.WriteLine($"Error opening website: {ex.Message}"); }
            };

            var exportButton = new Button { Text = "Export LCA Calculation", ToolTip = "Export the grid data to a CSV file." };
            exportButton.Click += OnExportToCsvClick;

            var buttonLayout = new StackLayout { Orientation = Orientation.Horizontal, Spacing = 5, Items = { btnStructure, exportButton } };
            var layout = new DynamicLayout { Spacing = new Size(6, 6) };
            layout.AddRow(buttonLayout);
            return layout;
        }


        private Control CreateDefinitionLayout()
        {
            var layout = new DynamicLayout { Spacing = new Size(6, 6) };

            var selectIcon = BytesToEtoBitmap(Resources.btn_selectObjects256, new Size(18, 18));
            var refreshIcon = BytesToEtoBitmap(Resources.btn_refreshList256, new Size(18, 18));

            var selectIfcButton = new Button { Image = selectIcon, ToolTip = "Select objects matching the specified IfcClass.", MinimumSize = Size.Empty };
            selectIfcButton.Click += OnSelectByIfcClassClick;
            var assignIfcButton = new Button { Image = BytesToEtoBitmap(Resources.btn_assignIfcClass256, new Size(18, 18)), MinimumSize = Size.Empty, ImagePosition = ButtonImagePosition.Left };
            assignIfcButton.Click += OnAssignIfcClassClick;
            var removeIfcButton = new Button { Image = BytesToEtoBitmap(Resources.btn_removeIfcClass256, new Size(18, 18)), MinimumSize = Size.Empty, ImagePosition = ButtonImagePosition.Left };
            removeIfcButton.Click += (s, e) => RemoveUserString(IfcClassKey, "IfcClass");
            var refreshIfcButton = new Button { Image = refreshIcon, ToolTip = "Refresh IFC Class list from source.", MinimumSize = Size.Empty };
            refreshIfcButton.Click += (s, e) => ReloadIfcClassList();

            var ifcClassLayout = new StackLayout { Orientation = Orientation.Horizontal, Spacing = 5, Items = { _ifcDropdown, _ifcSubclassDropdown } };
            var ifcButtonsLayout = new StackLayout { Orientation = Orientation.Horizontal, Spacing = 5, Items = { assignIfcButton, removeIfcButton, selectIfcButton, refreshIfcButton } };
            layout.AddRow(new Label { Text = "IfcClass:", ToolTip = IfcClassKey }, ifcClassLayout, ifcButtonsLayout);

            var selectMaterialButton = new Button { Image = selectIcon, ToolTip = "Select objects matching the specified Material.", MinimumSize = Size.Empty };
            selectMaterialButton.Click += OnSelectByMaterialClick;
            var assignMaterialButton = new Button { Image = BytesToEtoBitmap(Resources.btn_assignMaterial256, new Size(18, 18)), MinimumSize = Size.Empty, ImagePosition = ButtonImagePosition.Left };
            assignMaterialButton.Click += (s, e) => AssignUserString(MaterialKey, "Material", _materialDropdown.SelectedValue as ListItem);
            var removeMaterialButton = new Button { Image = BytesToEtoBitmap(Resources.btn_removeMaterial256, new Size(18, 18)), MinimumSize = Size.Empty, ImagePosition = ButtonImagePosition.Left };
            removeMaterialButton.Click += (s, e) => RemoveUserString(MaterialKey, "Material");
            var refreshMaterialButton = new Button { Image = refreshIcon, ToolTip = "Refresh Material list from source.", MinimumSize = Size.Empty };
            refreshMaterialButton.Click += (s, e) => ReloadMaterialList();

            var materialButtonsLayout = new StackLayout { Orientation = Orientation.Horizontal, Spacing = 5, Items = { assignMaterialButton, removeMaterialButton, selectMaterialButton, refreshMaterialButton } };
            layout.AddRow(new Label { Text = "Material:", ToolTip = MaterialKey }, _materialDropdown, materialButtonsLayout);

            _assignDefOnCreateCheckBox = new CheckBox { Text = "Assign On Creation", ToolTip = "If checked, automatically assigns the selected IfcClass and Material to newly created objects." };
            layout.AddRow(null, _assignDefOnCreateCheckBox);

            return layout;
        }


        private Control CreateCustomDefinitionLayout()
        {
            var layout = new DynamicLayout { Spacing = new Size(6, 6) };

            _customKeyDropdown = new ComboBox { Width = 90, AutoComplete = true };
            _customValueDropdown = new ComboBox { Width = 90, AutoComplete = true };

            var selectIcon = BytesToEtoBitmap(Resources.btn_selectObjects256, new Size(18, 18));
            var refreshIcon = BytesToEtoBitmap(Resources.btn_refreshList256, new Size(18, 18));

            var selectCustomButton = new Button { Image = selectIcon, ToolTip = "Select objects matching the specified custom key and/or value.", MinimumSize = Size.Empty };
            selectCustomButton.Click += OnSelectByCustomAttributeClick;
            var assignButton = new Button { Image = BytesToEtoBitmap(Resources.btn_assignCustom256, new Size(18, 18)), MinimumSize = Size.Empty, ImagePosition = ButtonImagePosition.Left };
            assignButton.Click += OnAssignCustomAttributeClick;
            var removeButton = new Button { Image = BytesToEtoBitmap(Resources.btn_removeCustom256, new Size(18, 18)), MinimumSize = Size.Empty, ImagePosition = ButtonImagePosition.Left };
            removeButton.Click += OnRemoveCustomAttributeClick;
            var refreshCustomButton = new Button { Image = refreshIcon, ToolTip = "Refresh Custom Attribute lists from source.", MinimumSize = Size.Empty };
            refreshCustomButton.Click += (s, e) => ReloadCustomAttributeList();

            var dropdownLayout = new StackLayout { Orientation = Orientation.Horizontal, Spacing = 5, Items = { _customKeyDropdown, _customValueDropdown } };
            var buttonsLayout = new StackLayout { Orientation = Orientation.Horizontal, Spacing = 5, Items = { assignButton, removeButton, selectCustomButton } };
            layout.AddRow(new Label { Text = "Custom:", ToolTip = "Assigns a single, replaceable custom attribute." }, dropdownLayout, buttonsLayout, refreshCustomButton);

            _assignCustomOnCreateCheckBox = new CheckBox { Text = "Assign On Creation", ToolTip = "If checked, automatically assigns the selected custom attribute to newly created objects." };
            layout.AddRow(null, _assignCustomOnCreateCheckBox);

            return layout;
        }


        private Control CreateAttributeGridLayout()
        {
            if (_userTextGridView.Columns.Count == 0)
            {
                _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.SerialNumber)), HeaderText = "SN", Editable = false, Width = 30 });
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

            var columnsViewButton = new Button { Image = BytesToEtoBitmap(Resources.btn_columnsView256, new Size(18, 18)), MinimumSize = Size.Empty, ImagePosition = ButtonImagePosition.Left };
            // --- MODIFICATION START: Save column settings after the dialog is closed ---
            columnsViewButton.Click += (s, e) =>
            {
                new ColumnVisibilityDialog(_userTextGridView).ShowModal(this);
                PluginSettingsManager.SaveColumnVisibility(_userTextGridView.Columns);
            };
            // --- MODIFICATION END ---
            var selectUnassignedButton = new Button { Image = BytesToEtoBitmap(Resources.btn_selectUnassigned256, new Size(18, 18)), ToolTip = "Select objects with no material assigned.", MinimumSize = Size.Empty };
            selectUnassignedButton.Click += OnSelectUnassignedClick;

            var viewOptionsLayout = new StackLayout { Orientation = Orientation.Horizontal, Spacing = 10, Items = { _showAllObjectsCheckBox, _showUnassignedCheckBox } };
            var groupingOptionsLayout = new StackLayout { Orientation = Orientation.Horizontal, Spacing = 10, Items = { _groupByClassCheckBox, _groupByMaterialCheckBox, _aggregateCheckBox, null, selectUnassignedButton, columnsViewButton } };

            var layout = new DynamicLayout { Spacing = new Size(6, 6) };
            layout.AddRow(viewOptionsLayout);
            layout.AddRow(groupingOptionsLayout);
            layout.AddRow(new Scrollable { Content = _userTextGridView, Width = 450, Height = 220 });
            layout.AddRow(_totalLcaLabel);
            return layout;
        }


        private Control CreateObjectUserTextLayout()
        {
            var layout = new DynamicLayout { Spacing = new Size(6, 6) };

            _objectUserTextGridView.Columns.Add(new GridColumn { HeaderText = "Key", DataCell = new TextBoxCell(nameof(ObjectUserTextEntry.Key)), Editable = false, Width = 150 });
            _objectUserTextGridView.Columns.Add(new GridColumn { HeaderText = "Value", DataCell = new TextBoxCell(nameof(ObjectUserTextEntry.Value)), Editable = false });

            layout.AddRow(new Scrollable { Content = _objectUserTextGridView });
            return layout;
        }


        private Control CreateDisplayOptionsLayout()
        {
            _displayIfcClassCheckBox = new CheckBox { Text = "Display IFC Class" };
            _ifcLeaderLengthTextBox = new TextBox { Text = _ifcClassDisplayConduit.LeaderLength.ToString(), Width = 40 };
            _ifcLeaderAngleTextBox = new TextBox { Text = _ifcClassDisplayConduit.LeaderAngle.ToString(), Width = 40 };
            var applyIfcButton = new Button { Text = "Apply" };
            applyIfcButton.Click += OnApplyIfcDisplaySettingsClick;

            _displayMaterialCheckBox = new CheckBox { Text = "Display Material" };
            _materialLeaderLengthTextBox = new TextBox { Text = _materialDisplayConduit.LeaderLength.ToString(), Width = 40 };
            _materialLeaderAngleTextBox = new TextBox { Text = _materialDisplayConduit.LeaderAngle.ToString(), Width = 40 };
            var applyMaterialButton = new Button { Text = "Apply" };
            applyMaterialButton.Click += OnApplyMaterialDisplaySettingsClick;

            var layout = new DynamicLayout { Spacing = new Size(6, 6) };

            layout.AddRow(_displayIfcClassCheckBox);
            layout.AddRow(
                new Label { Text = "Length mult.:" }, _ifcLeaderLengthTextBox,
                new Label { Text = "Angle:" }, _ifcLeaderAngleTextBox,
                null,
                applyIfcButton
            );
            layout.AddRow(_displayMaterialCheckBox);
            layout.AddRow(
                new Label { Text = "Length mult.:" }, _materialLeaderLengthTextBox,
                new Label { Text = "Angle:" }, _materialLeaderAngleTextBox,
                null,
                applyMaterialButton
            );

            return layout;
        }


        private Control CreateDocumentUserTextLayout()
        {
            var layout = new DynamicLayout { Spacing = new Size(6, 6) };

            _docUserTextGridView.Columns.Add(new GridColumn { HeaderText = "Key", DataCell = new TextBoxCell(nameof(DocumentUserTextEntry.Key)), Editable = false, Width = 150 });
            _docUserTextGridView.Columns.Add(new GridColumn { HeaderText = "Value", DataCell = new TextBoxCell(nameof(DocumentUserTextEntry.Value)), Editable = false });

            var refreshButton = new Button { Text = "Refresh User Text" };
            refreshButton.Click += (s, e) => UpdatePanelData();

            layout.AddRow(new Scrollable { Content = _docUserTextGridView });
            layout.AddRow(refreshButton);
            return layout;
        }

        #endregion

        #region Rhino & UI Event Handlers


        private void RegisterEventHandlers()
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

            _ifcDropdown.SelectedValueChanged += OnPrimaryIfcClassChanged;

            // --- MODIFICATION START: Add setting-saving logic to event handlers ---
            _showAllObjectsCheckBox.CheckedChanged += (s, e) => {
                PluginSettingsManager.SetBool(SettingKeys.ShowAllObjects, _showAllObjectsCheckBox.Checked ?? false);
                UpdatePanelData();
            };
            _showUnassignedCheckBox.CheckedChanged += (s, e) => {
                PluginSettingsManager.SetBool(SettingKeys.ShowUnassigned, _showUnassignedCheckBox.Checked ?? false);
                UpdatePanelData();
            };
            _aggregateCheckBox.CheckedChanged += (s, e) => {
                PluginSettingsManager.SetBool(SettingKeys.AggregateResults, _aggregateCheckBox.Checked ?? false);
                OnAggregateChanged(s, e);
            };
            _groupByMaterialCheckBox.CheckedChanged += (s, e) => {
                PluginSettingsManager.SetBool(SettingKeys.GroupByMaterial, _groupByMaterialCheckBox.Checked ?? false);
                OnGroupByMaterialChanged(s, e);
            };
            _groupByClassCheckBox.CheckedChanged += (s, e) => {
                PluginSettingsManager.SetBool(SettingKeys.GroupByClass, _groupByClassCheckBox.Checked ?? false);
                OnGroupByClassChanged(s, e);
            };
            _displayIfcClassCheckBox.CheckedChanged += (s, e) => {
                PluginSettingsManager.SetBool(SettingKeys.DisplayIfcClass, _displayIfcClassCheckBox.Checked ?? false);
                OnDisplayAttributeTextCheckedChanged(s, e);
            };
            _displayMaterialCheckBox.CheckedChanged += (s, e) => {
                PluginSettingsManager.SetBool(SettingKeys.DisplayMaterial, _displayMaterialCheckBox.Checked ?? false);
                OnDisplayAttributeTextCheckedChanged(s, e);
            };
            _assignDefOnCreateCheckBox.CheckedChanged += (s, e) => {
                PluginSettingsManager.SetBool(SettingKeys.AssignDefOnCreate, _assignDefOnCreateCheckBox.Checked ?? false);
            };
            _assignCustomOnCreateCheckBox.CheckedChanged += (s, e) => {
                PluginSettingsManager.SetBool(SettingKeys.AssignCustomOnCreate, _assignCustomOnCreateCheckBox.Checked ?? false);
            };
            // --- MODIFICATION END ---

            _userTextGridView.SelectionChanged += OnGridSelectionChanged;
            _userTextGridView.CellEdited += OnGridCellEdited;
            _userTextGridView.MouseEnter += (s, e) => _isMouseOverGrid = true;
            _userTextGridView.MouseLeave += (s, e) => _isMouseOverGrid = false;
        }

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
                        if (newAttributes.SetUserString(MaterialKey, materialValue))
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
        private void OnExportToCsvClick(object sender, EventArgs e)
        {
            if (_userTextGridView.DataStore == null || !_userTextGridView.DataStore.Any())
            {
                RhinoApp.WriteLine("No data available to export.");
                return;
            }

            var data = _userTextGridView.DataStore.ToList();
            var dialog = new SaveFileDialog
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
                .Where(obj => obj.Attributes.GetUserString(MaterialKey) == materialTarget)
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
        private void OnAssignIfcClassClick(object sender, EventArgs e)
        {
            string primaryClass = (_ifcDropdown.SelectedValue as ListItem)?.Text;
            string secondaryClass = (_ifcSubclassDropdown.SelectedValue as ListItem)?.Text;

            if (string.IsNullOrWhiteSpace(primaryClass))
            {
                RhinoApp.WriteLine("Please select a primary IFC class from the dropdown first.");
                return;
            }

            string combinedClass = !string.IsNullOrWhiteSpace(secondaryClass) ? $"{primaryClass}, {secondaryClass}" : primaryClass;
            AssignUserString(IfcClassKey, "IfcClass", combinedClass);
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
        private void OnPrimaryIfcClassChanged(object sender, EventArgs e) => UpdateSubclassDropdown();
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
                    string.IsNullOrWhiteSpace(o.Attributes.GetUserString(MaterialKey)))
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
        private void OnDisplayAttributeTextCheckedChanged(object sender, EventArgs e)
        {
            UpdateAllConduits();
            RhinoDoc.ActiveDoc?.Views.Redraw();
        }

        private void OnApplyIfcDisplaySettingsClick(object sender, EventArgs e)
        {
            if (ValidateAndApplyConduitSettings(_ifcLeaderLengthTextBox, _ifcLeaderAngleTextBox, _ifcClassDisplayConduit))
            {
                // --- MODIFICATION START: Save settings on successful apply ---
                PluginSettingsManager.SetString(SettingKeys.IfcLeaderLength, _ifcLeaderLengthTextBox.Text);
                PluginSettingsManager.SetString(SettingKeys.IfcLeaderAngle, _ifcLeaderAngleTextBox.Text);
                // --- MODIFICATION END ---
                RhinoDoc.ActiveDoc?.Views.Redraw();
            }
        }

        private void OnApplyMaterialDisplaySettingsClick(object sender, EventArgs e)
        {
            if (ValidateAndApplyConduitSettings(_materialLeaderLengthTextBox, _materialLeaderAngleTextBox, _materialDisplayConduit))
            {
                // --- MODIFICATION START: Save settings on successful apply ---
                PluginSettingsManager.SetString(SettingKeys.MaterialLeaderLength, _materialLeaderLengthTextBox.Text);
                PluginSettingsManager.SetString(SettingKeys.MaterialLeaderAngle, _materialLeaderAngleTextBox.Text);
                // --- MODIFICATION END ---
                RhinoDoc.ActiveDoc?.Views.Redraw();
            }
        }

        #endregion

        #region Data Loading & Panel Updates

        // --- NEW METHOD START ---
        /// <summary>
        /// Loads all panel settings from the persistent settings store and applies them to the UI controls.
        /// </summary>
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

            // Load leader settings from conduit properties which were initialized with saved values
            _ifcLeaderLengthTextBox.Text = _ifcClassDisplayConduit.LeaderLength.ToString();
            _ifcLeaderAngleTextBox.Text = _ifcClassDisplayConduit.LeaderAngle.ToString();
            _materialLeaderLengthTextBox.Text = _materialDisplayConduit.LeaderLength.ToString();
            _materialLeaderAngleTextBox.Text = _materialDisplayConduit.LeaderAngle.ToString();

            // Load grid column visibility
            PluginSettingsManager.LoadColumnVisibility(_userTextGridView.Columns);
        }
        // --- NEW METHOD END ---

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

            var objectsToProcess = (_showAllObjectsCheckBox.Checked == true
              ? doc.Objects.Where(obj => obj.IsSelectable(true, false, false, true) && obj.Attributes.Visible && doc.Layers[obj.Attributes.LayerIndex].IsVisible)
              : doc.Objects.GetSelectedObjects(false, false)).ToList();

            if (!objectsToProcess.Any())
            {
                _userTextGridView.DataStore = null;
                _totalLcaLabel.Text = "Total LCA: 0.00";
                return;
            }

            var processedObjects = objectsToProcess.Select(o =>
            {
                TryComputeQuantity(o.Geometry, doc, out double qty, out string unit, out string qtyType);

                var multiplierString = o.Attributes.GetUserString(QuantityMultiplierKey);
                double multiplier = 1.0;
                if (!string.IsNullOrEmpty(multiplierString))
                {
                    if (double.TryParse(multiplierString, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedMultiplier) && parsedMultiplier > 0)
                    {
                        multiplier = parsedMultiplier;
                    }
                }

                return new
                {
                    RhinoObject = o,
                    Material = o.Attributes.GetUserString(MaterialKey) ?? "N/A",
                    IfcClass = o.Attributes.GetUserString(IfcClassKey) ?? "N/A",
                    Quantity = qty,
                    QuantityUnit = unit,
                    QuantityType = qtyType,
                    QuantityMultiplier = multiplier
                };
            }).ToList();

            List<UserTextEntry> gridData;
            if (_aggregateCheckBox.Checked == true)
            {
                var aggregatedData = processedObjects
                  .GroupBy(p => new { p.Material, p.IfcClass, p.QuantityType })
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
                          ObjectIds = g.Select(p => p.RhinoObject.Id).ToList()
                      };
                  });
                gridData = SortAndFormatData(aggregatedData.ToList());
            }
            else
            {
                var data = processedObjects.Select(p =>
                {
                    double lca = 0.0;
                    _materialLcaData.TryGetValue(p.Material, out MaterialData materialData);
                    if (materialData != null && p.QuantityType != "N/A") lca = (p.Quantity * p.QuantityMultiplier) * materialData.Lca;

                    return new UserTextEntry
                    {
                        Value = p.Material,
                        IfcClass = p.IfcClass,
                        ReferenceQuantity = materialData?.ReferenceQuantity ?? 0,
                        ReferenceUnit = materialData?.ReferenceUnit ?? "",
                        ReferenceLca = materialData?.Lca ?? 0,
                        Count = 1,
                        Quantity = p.Quantity,
                        QuantityUnit = p.QuantityUnit,
                        QuantityType = p.QuantityType,
                        QuantityMultiplier = p.QuantityMultiplier,
                        Lca = lca,
                        ObjectIds = new List<Guid> { p.RhinoObject.Id }
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
                    var material = rhinoObject.Attributes.GetUserString(MaterialKey);
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
            CsvReader.ReadCustomAttributeListsFromResource("MyProject.Resources.Data.customAttribute.csv", out _customAttributeKeys, out _customAttributeValues);
            PopulateCustomKeyDropdown();
            PopulateCustomValueDropdown();
        }
        private void PopulateCustomKeyDropdown()
        {
            _customKeyDropdown.Items.Clear();
            _customKeyDropdown.Items.Add(new ListItem { Text = "" });
            if (_customAttributeKeys != null)
            {
                foreach (var key in _customAttributeKeys)
                    _customKeyDropdown.Items.Add(new ListItem { Text = key });
            }
            _customKeyDropdown.SelectedIndex = 0;
        }
        private void PopulateCustomValueDropdown()
        {
            _customValueDropdown.Items.Clear();
            _customValueDropdown.Items.Add(new ListItem { Text = "" });
            if (_customAttributeValues != null)
            {
                foreach (var value in _customAttributeValues)
                    _customValueDropdown.Items.Add(new ListItem { Text = value });
            }
            _customValueDropdown.SelectedIndex = 0;
        }
        private void PopulateMaterialDropdown()
        {
            var selectedText = _materialDropdown.SelectedValue is ListItem selectedItem ? selectedItem.Text : null;
            _materialDropdown.Items.Clear();
            _materialDropdown.Items.Add(new ListItem { Text = "" });
            if (_materialLcaData != null)
            {
                foreach (var materialName in _materialLcaData.Keys)
                    _materialDropdown.Items.Add(new ListItem { Text = materialName });
            }
            if (selectedText != null)
            {
                var itemToRestore = _materialDropdown.Items.FirstOrDefault(i => i.Text == selectedText);
                if (itemToRestore != null) _materialDropdown.SelectedValue = itemToRestore;
                else _materialDropdown.SelectedIndex = 0;
            }
            else
            {
                _materialDropdown.SelectedIndex = 0;
            }
        }
        private void PopulatePrimaryIfcDropdown()
        {
            var selectedText = _ifcDropdown.SelectedValue is ListItem selectedItem ? selectedItem.Text : null;
            _ifcDropdown.Items.Clear();
            _ifcDropdown.Items.Add(new ListItem { Text = "" });
            foreach (var primaryClass in _ifcClasses.Keys)
            {
                _ifcDropdown.Items.Add(new ListItem { Text = primaryClass });
            }
            if (selectedText != null)
            {
                var itemToRestore = _ifcDropdown.Items.FirstOrDefault(i => i.Text == selectedText);
                if (itemToRestore != null) _ifcDropdown.SelectedValue = itemToRestore;
            }
            else
            {
                _ifcDropdown.SelectedIndex = 0;
            }
            UpdateSubclassDropdown();
        }
        private void UpdateSubclassDropdown()
        {
            _ifcSubclassDropdown.Items.Clear();
            _ifcSubclassDropdown.Enabled = false;

            if (_ifcDropdown.SelectedValue is ListItem selectedItem && !string.IsNullOrEmpty(selectedItem.Text))
            {
                string primaryClass = selectedItem.Text;
                if (_ifcClasses.TryGetValue(primaryClass, out List<string> subclasses) && subclasses.Any())
                {
                    _ifcSubclassDropdown.Items.Add(new ListItem { Text = "" });
                    foreach (var subclass in subclasses)
                    {
                        _ifcSubclassDropdown.Items.Add(new ListItem { Text = subclass });
                    }
                    _ifcSubclassDropdown.SelectedIndex = 0;
                    _ifcSubclassDropdown.Enabled = true;
                }
            }
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

            int selectionCount = selectedObjects.Count;
            string objectText = selectionCount == 1 ? "Object" : "Objects";
            _selectedObjectAttributesHeaderLabel.Text = $"Selected Object(s) Attributes ({selectionCount} {objectText} Selected)";

            if (!selectedObjects.Any())
            {
                _objectUserTextGridView.DataStore = null;
                return;
            }

            var allAttributes = new Dictionary<string, HashSet<string>>();

            foreach (var obj in selectedObjects)
            {
                var userStrings = obj.Attributes.GetUserStrings();
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

            _objectUserTextGridView.DataStore = data;
        }

        #endregion

        #region Core Logic & Helpers

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
                // No need to unregister event handlers for settings as they are tied to the control's lifetime.
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
        private static void AssignUserString(string key, string friendlyName, ListItem selectedItem)
        {
            var value = selectedItem?.Text;
            if (string.IsNullOrWhiteSpace(value))
            {
                RhinoApp.WriteLine($"Please select a valid {friendlyName} from the dropdown first.");
                return;
            }
            AssignUserString(key, friendlyName, value);
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