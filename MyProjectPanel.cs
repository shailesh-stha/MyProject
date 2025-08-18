// MyProjectPanel.cs
using Eto.Drawing;
using Eto.Forms;
using Rhino;
using Rhino.DocObjects;
using Rhino.Geometry;
using Rhino.Input.Custom;
using Rhino.UI.Controls;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Reflection;

namespace MyProject
{
    [Guid("01D22D46-3C89-43A5-A92B-A1D6519A4244")]
    public class MyProjectPanel : Panel
    {
        public static string SelectedIfcClass { get; private set; }
        private readonly List<string> _ifcClasses;

        public static string SelectedMaterial { get; private set; }
        private readonly Dictionary<string, MaterialData> _materialLcaData;

        private Label _totalLcaLabel;

        private readonly CheckBox _showAllObjectsCheckBox;
        private readonly CheckBox _showUnassignedCheckBox;
        private readonly CheckBox _groupByMaterialCheckBox;
        private readonly CheckBox _groupByClassCheckBox;
        private readonly CheckBox _aggregateCheckBox;

        private bool _isSyncingSelection = false;
        private bool _needsRefresh = false;

        private readonly Dictionary<Guid, double> _quantityMultipliers = new Dictionary<Guid, double>();

        private class UserTextEntry
        {
            public string SerialNumber { get; set; }
            public string IfcClass { get; set; }
            public string Value { get; set; } // Material Name
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

        private readonly GridView<UserTextEntry> _userTextGridView;
        private readonly Button _btnColumns;
        private readonly GridColumn _qtyMultiplierColumn;

        public MyProjectPanel(uint documentSerialNumber)
        {
            _materialLcaData = CsvReader.ReadMaterialLcaDataFromResource("MyProject.Resources.materialListWithUnits.csv");
            _ifcClasses = new List<string> { "Wall", "Slab", "Beam", "Column", "Foundation", "Roof", "Stair", "Ramp", "Door", "Window", "Railing", "Covering" };

            Styles.Add<Label>("bold_label", label => label.Font = SystemFonts.Bold());

            #region Upper Layout (General Buttons)
            var btnStructure = new Button { Text = "str-ucture" };
            btnStructure.Click += (s, e) =>
            {
                var url = "http://www.str-ucture.com";
                try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
                catch (Exception ex) { RhinoApp.WriteLine($"Error opening website: {ex.Message}"); }
            };

            var btnSelectUnassigned = new Button { Text = "Select Unassigned" };
            btnSelectUnassigned.Click += OnSelectUnassignedClick;

            _btnColumns = new Button { Text = "Columns" };
            _btnColumns.Click += OnColumnsButtonClick;

            var buttonLayout = new StackLayout
            {
                Orientation = Orientation.Horizontal,
                Spacing = 5,
                Items = { btnSelectUnassigned, _btnColumns, btnStructure }
            };

            var upperLayout = new DynamicLayout { Spacing = new Size(6, 6), Padding = new Padding(10) };
            upperLayout.AddRow(new Label { Text = "Buttons for testing", Style = "bold_label" });
            upperLayout.AddRow(buttonLayout);
            upperLayout.Add(null);
            #endregion

            #region IFC Class Layout
            // MODIFIED: Added AutoComplete = true to enable filtering
            var ifcDropdown = new ComboBox { Width = 180, AutoComplete = true };
            ifcDropdown.Items.Add(new ListItem { Text = "" });
            foreach (var cls in _ifcClasses) ifcDropdown.Items.Add(new ListItem { Text = cls });

            ifcDropdown.SelectedIndexChanged += (s, e) =>
            {
                var li = ifcDropdown.SelectedValue as ListItem;
                SelectedIfcClass = li?.Text;
            };

            var ifcAssignButton = new Button { Text = "Assign" };
            ifcAssignButton.Click += OnIfcAssignButtonClick;

            var ifcRemoveButton = new Button { Text = "Remove" };
            ifcRemoveButton.Click += OnIfcRemoveButtonClick;

            var ifcButtonsLayout = new StackLayout
            {
                Orientation = Orientation.Horizontal,
                Spacing = 5,
                Items = { ifcAssignButton, ifcRemoveButton }
            };

            var ifcLayout = new DynamicLayout { Spacing = new Size(6, 6), Padding = new Padding(10) };
            ifcLayout.AddRow(new Divider());
            ifcLayout.AddRow(new Label { Text = "IFC Class Definition", Style = "bold_label" });
            ifcLayout.AddRow(new Label { Text = "Select IfcClass:" }, ifcDropdown, ifcButtonsLayout);
            ifcLayout.Add(null);
            #endregion

            #region Material Definition Layout
            // MODIFIED: Added AutoComplete = true to enable filtering
            var materialDropdown = new ComboBox { Width = 180, AutoComplete = true };
            materialDropdown.Items.Add(new ListItem { Text = "" });
            foreach (var m in _materialLcaData.Keys) materialDropdown.Items.Add(new ListItem { Text = m });

            materialDropdown.SelectedIndexChanged += (s, e) =>
            {
                var li = materialDropdown.SelectedValue as ListItem;
                var val = li?.Text;
                SelectedMaterial = string.IsNullOrWhiteSpace(val) ? null : val;
            };

            var assignButton = new Button { Text = "Assign" };
            assignButton.Click += OnAssignButtonClick;

            var removeButton = new Button { Text = "Remove" };
            removeButton.Click += OnRemoveButtonClick;

            var materialButtonsLayout = new StackLayout
            {
                Orientation = Orientation.Horizontal,
                Spacing = 5,
                Items = { assignButton, removeButton }
            };

            var middleLayout = new DynamicLayout { Spacing = new Size(6, 6), Padding = new Padding(10) };
            middleLayout.AddRow(new Divider());
            middleLayout.AddRow(new Label { Text = "Material Definition", Style = "bold_label" });
            middleLayout.AddRow(new Label { Text = "Select Material:" }, materialDropdown, materialButtonsLayout);
            middleLayout.Add(null);
            #endregion

            #region GridView and Data Layout
            _userTextGridView = new GridView<UserTextEntry>
            {
                ShowHeader = true,
                AllowMultipleSelection = true
            };

            var columns = new List<GridColumn>
            {
                // First 3 columns use default (left) alignment
                new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.SerialNumber)), HeaderText = "SN", Editable = false, Width = 30 },
                new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.IfcClass)), HeaderText = "IfcClass", Editable = false, Width = 60 },
                new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.Value)), HeaderText = "Material (Name)", Editable = false, Width = 150 },
                
                // The following columns now have their text aligned to the right.
                new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.DisplayReference)) { TextAlignment = TextAlignment.Right }, HeaderText = "Ref. Qty.", Editable = false, Width = 70 },
                new GridColumn { DataCell = new TextBoxCell { Binding = Binding.Property<UserTextEntry, string>(entry => $"{entry.ReferenceLca:0.00}"), TextAlignment = TextAlignment.Right }, HeaderText = "Ref. LCA", Editable = false, Width = 70 },
                new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.Count)) { TextAlignment = TextAlignment.Right }, HeaderText = "Count", Editable = false, Width = 50 },
                new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.DisplayQuantity)) { TextAlignment = TextAlignment.Right }, HeaderText = "Qty. (Rh)", Editable = false, Width = 75 }
            };

            // These columns also have their text aligned to the right.
            _qtyMultiplierColumn = new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.QuantityMultiplier)) { TextAlignment = TextAlignment.Right }, HeaderText = "Qty. Multiplier", Editable = true, Width = 90 };
            var qtyTotalColumn = new GridColumn { DataCell = new TextBoxCell { Binding = Binding.Property<UserTextEntry, string>(entry => $"{entry.QuantityTotal:0.00}"), TextAlignment = TextAlignment.Right }, HeaderText = "Qty. Total", Editable = false, Width = 75 };

            columns.Add(_qtyMultiplierColumn);
            columns.Add(qtyTotalColumn);

            // This column also has its text aligned to the right.
            columns.Add(new GridColumn { DataCell = new TextBoxCell { Binding = Binding.Property<UserTextEntry, string>(entry => $"{entry.Lca:0.00}"), TextAlignment = TextAlignment.Right }, HeaderText = "LCA (kgCO2 eq)", Editable = false, Width = 100 });

            // Add all defined columns to the grid view.
            foreach (var col in columns)
            {
                _userTextGridView.Columns.Add(col);
            }

            _userTextGridView.Width = 550;

            _userTextGridView.SelectionChanged += OnGridSelectionChanged;
            _userTextGridView.CellEdited += OnGridCellEdited;


            _showAllObjectsCheckBox = new CheckBox { Text = "Show All/Selected", Checked = true };
            _showAllObjectsCheckBox.CheckedChanged += (s, e) => UpdatePanelData();

            _showUnassignedCheckBox = new CheckBox { Text = "Show/Hide Unassigned", Checked = true };
            _showUnassignedCheckBox.CheckedChanged += (s, e) => UpdatePanelData();

            _groupByMaterialCheckBox = new CheckBox { Text = "Group by Material", Checked = false };
            _groupByClassCheckBox = new CheckBox { Text = "Group by Class", Checked = true };
            _aggregateCheckBox = new CheckBox { Text = "Aggregate", Checked = false };

            _aggregateCheckBox.CheckedChanged += (s, e) =>
            {
                _qtyMultiplierColumn.Visible = !_aggregateCheckBox.Checked ?? true;
                UpdatePanelData();
            };

            _qtyMultiplierColumn.Visible = !_aggregateCheckBox.Checked ?? true;


            _groupByMaterialCheckBox.CheckedChanged += OnGroupByMaterialChanged;
            _groupByClassCheckBox.CheckedChanged += (s, e) =>
            {
                if (_groupByClassCheckBox.Checked == true) _groupByMaterialCheckBox.Checked = false;
                UpdatePanelData();
            };

            var scrollableGrid = new Scrollable
            {
                Content = _userTextGridView,
                Height = 280
            };

            var viewOptionsLayout = new StackLayout
            {
                Orientation = Orientation.Horizontal,
                Spacing = 10,
                Items = { _showAllObjectsCheckBox, _showUnassignedCheckBox }
            };

            var groupingOptionsLayout = new StackLayout
            {
                Orientation = Orientation.Horizontal,
                Spacing = 10,
                Items = { _groupByClassCheckBox, _groupByMaterialCheckBox, _aggregateCheckBox }
            };
            var checkBoxLayout = new StackLayout
            {
                Orientation = Orientation.Vertical,
                Spacing = 5,
                Items = { viewOptionsLayout, groupingOptionsLayout }
            };

            _totalLcaLabel = new Label { Text = "Total LCA: 0.00", Style = "bold_label" };

            var gridHeaderLayout = new DynamicLayout();
            gridHeaderLayout.AddRow(new Label { Text = "Attribute User Text", Style = "bold_label" }, null, _btnColumns);

            var userTextLayout = new DynamicLayout { Spacing = new Size(6, 6), Padding = new Padding(10) };
            userTextLayout.AddRow(new Divider());
            userTextLayout.AddRow(gridHeaderLayout);
            userTextLayout.AddRow(checkBoxLayout);
            userTextLayout.AddRow(scrollableGrid);
            userTextLayout.AddRow(_totalLcaLabel, null);
            userTextLayout.Add(null);
            #endregion

            #region Final Panel Assembly
            var lowerContainer = new StackLayout { Orientation = Orientation.Vertical, Spacing = 0, Items = { ifcLayout, middleLayout, userTextLayout } };
            var splitter = new Splitter { Orientation = Orientation.Vertical, Panel1 = upperLayout, Panel2 = lowerContainer, FixedPanel = SplitterFixedPanel.Panel1, Position = 70 };

            Content = splitter;
            MinimumSize = new Size(400, 5500);

            RhinoDoc.SelectObjects += OnSelectionChanged;
            RhinoDoc.DeselectObjects += OnSelectionChanged;
            RhinoDoc.DeselectAllObjects += OnDeselectAllObjects;
            RhinoDoc.AddRhinoObject += OnDatabaseChanged;
            RhinoDoc.DeleteRhinoObject += OnDatabaseChanged;
            RhinoDoc.ReplaceRhinoObject += OnReplaceObject;
            RhinoDoc.UserStringChanged += OnUserStringChanged;
            RhinoDoc.EndOpenDocument += OnDocumentChanged;
            RhinoDoc.NewDocument += OnDocumentChanged;
            RhinoApp.Idle += OnRhinoIdle;

            UpdatePanelData();
            #endregion
        }

        #region Event Handlers

        private void OnGridCellEdited(object sender, GridViewCellEventArgs e)
        {
            if (e.Item is UserTextEntry entry && entry.ObjectIds.Any())
            {
                var objectId = entry.ObjectIds.First();
                _quantityMultipliers[objectId] = entry.QuantityMultiplier;
                UpdatePanelData();
            }
        }

        private void OnColumnsButtonClick(object sender, EventArgs e)
        {
            var dialog = new ColumnVisibilityDialog(_userTextGridView);
            dialog.ShowModal(this);
        }

        private void OnGroupByMaterialChanged(object sender, EventArgs e)
        {
            if (_groupByMaterialCheckBox.Checked == true)
            {
                _groupByClassCheckBox.Checked = false;

                int ifcClassIndex = -1;
                int materialNameIndex = -1;

                for (int i = 0; i < _userTextGridView.Columns.Count; i++)
                {
                    if (_userTextGridView.Columns[i].HeaderText == "IfcClass") ifcClassIndex = i;
                    if (_userTextGridView.Columns[i].HeaderText == "Material (Name)") materialNameIndex = i;
                }

                if (ifcClassIndex != -1 && materialNameIndex != -1 && ifcClassIndex < materialNameIndex)
                {
                    var ifcColumn = _userTextGridView.Columns[ifcClassIndex];
                    _userTextGridView.Columns.RemoveAt(ifcClassIndex);
                    _userTextGridView.Columns.Insert(materialNameIndex, ifcColumn);
                }
            }
            else
            {
                int ifcClassIndex = -1;
                int materialNameIndex = -1;

                for (int i = 0; i < _userTextGridView.Columns.Count; i++)
                {
                    if (_userTextGridView.Columns[i].HeaderText == "IfcClass") ifcClassIndex = i;
                    if (_userTextGridView.Columns[i].HeaderText == "Material (Name)") materialNameIndex = i;
                }

                if (ifcClassIndex != -1 && materialNameIndex != -1 && ifcClassIndex > materialNameIndex)
                {
                    var ifcColumn = _userTextGridView.Columns[ifcClassIndex];
                    _userTextGridView.Columns.RemoveAt(ifcClassIndex);
                    _userTextGridView.Columns.Insert(materialNameIndex, ifcColumn);
                }
            }

            UpdatePanelData();
        }

        private void OnRhinoIdle(object sender, EventArgs e)
        {
            if (_needsRefresh)
            {
                _needsRefresh = false;
                UpdatePanelData();
            }
        }

        private void OnDocumentChanged(object sender, EventArgs e)
        {
            _quantityMultipliers.Clear();
            _needsRefresh = true;
        }

        private void OnIfcAssignButtonClick(object sender, EventArgs e)
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;
            if (string.IsNullOrWhiteSpace(SelectedIfcClass))
            {
                RhinoApp.WriteLine("Please select an IFC class from the dropdown first.");
                return;
            }
            var selectedObjects = doc.Objects.GetSelectedObjects(false, false).ToList();
            if (selectedObjects.Count == 0)
            {
                RhinoApp.WriteLine("Please select one or more objects in the viewport.");
                return;
            }

            int updatedCount = 0;
            foreach (var rhinoObject in selectedObjects)
            {
                if (rhinoObject == null) continue;
                var newAttributes = rhinoObject.Attributes.Duplicate();
                newAttributes.SetUserString("str-ifcclass", SelectedIfcClass);
                if (doc.Objects.ModifyAttributes(rhinoObject.Id, newAttributes, false)) updatedCount++;
            }
            doc.Views.Redraw();
            RhinoApp.WriteLine($"Successfully assigned '{SelectedIfcClass}' to {updatedCount} object(s).");
        }

        private void OnIfcRemoveButtonClick(object sender, EventArgs e)
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;

            var selectedObjects = doc.Objects.GetSelectedObjects(false, false).ToList();
            if (selectedObjects.Count == 0)
            {
                RhinoApp.WriteLine("Please select one or more objects to remove the IFC class from.");
                return;
            }

            int updatedCount = 0;
            foreach (var rhinoObject in selectedObjects)
            {
                if (rhinoObject?.Attributes.GetUserString("str-ifcclass") != null)
                {
                    var newAttributes = rhinoObject.Attributes.Duplicate();
                    newAttributes.DeleteUserString("str-ifcclass");
                    if (doc.Objects.ModifyAttributes(rhinoObject.Id, newAttributes, false)) updatedCount++;
                }
            }

            if (updatedCount > 0)
            {
                doc.Views.Redraw();
                RhinoApp.WriteLine($"Successfully removed IFC class definition from {updatedCount} object(s).");
            }
            else
            {
                RhinoApp.WriteLine("Selected objects do not have an IFC class definition to remove.");
            }
        }

        private void OnRemoveButtonClick(object sender, EventArgs e)
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;

            var selectedObjects = doc.Objects.GetSelectedObjects(false, false).ToList();
            if (selectedObjects.Count == 0)
            {
                RhinoApp.WriteLine("Please select one or more objects in the viewport to remove material definition.");
                return;
            }

            int updatedCount = 0;
            foreach (var rhinoObject in selectedObjects)
            {
                if (rhinoObject?.Attributes.GetUserString("str-material") != null)
                {
                    var newAttributes = rhinoObject.Attributes.Duplicate();
                    newAttributes.DeleteUserString("str-material");
                    if (doc.Objects.ModifyAttributes(rhinoObject.Id, newAttributes, false)) updatedCount++;
                }
            }

            if (updatedCount > 0)
            {
                doc.Views.Redraw();
                RhinoApp.WriteLine($"Successfully removed material definition from {updatedCount} object(s).");
            }
            else
            {
                RhinoApp.WriteLine("Selected objects do not have a material definition to remove.");
            }
        }

        private void OnAssignButtonClick(object sender, EventArgs e)
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;
            if (string.IsNullOrWhiteSpace(SelectedMaterial))
            {
                RhinoApp.WriteLine("Please select a material from the dropdown first.");
                return;
            }
            var selectedObjects = doc.Objects.GetSelectedObjects(false, false).ToList();
            if (selectedObjects.Count == 0)
            {
                RhinoApp.WriteLine("Please select one or more objects in the viewport.");
                return;
            }
            int updatedCount = 0;
            foreach (var rhinoObject in selectedObjects)
            {
                if (rhinoObject == null) continue;
                var newAttributes = rhinoObject.Attributes.Duplicate();
                newAttributes.SetUserString("str-material", SelectedMaterial);
                if (doc.Objects.ModifyAttributes(rhinoObject.Id, newAttributes, false)) updatedCount++;
            }
            doc.Views.Redraw();
            RhinoApp.WriteLine($"Successfully assigned '{SelectedMaterial}' to {updatedCount} object(s).");
        }

        private void OnGridSelectionChanged(object sender, EventArgs e)
        {
            if (_isSyncingSelection) return;

            try
            {
                _isSyncingSelection = true;
                var doc = RhinoDoc.ActiveDoc;
                if (doc == null || _userTextGridView.DataStore == null) return;
                var idsThatShouldBeSelected = _userTextGridView.SelectedItems.SelectMany(entry => entry.ObjectIds).ToHashSet();
                var allIdsInGrid = _userTextGridView.DataStore.SelectMany(entry => entry.ObjectIds).Distinct();
                foreach (var objId in allIdsInGrid)
                {
                    doc.Objects.Select(objId, idsThatShouldBeSelected.Contains(objId));
                }
                doc.Views.Redraw();
            }
            finally
            {
                _isSyncingSelection = false;
            }
        }

        private void OnSelectUnassignedClick(object sender, EventArgs e)
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;

            var idsToSelect = new List<Guid>();
            foreach (var rhinoObject in doc.Objects.Where(o => o.IsSelectable(true, false, false, true)))
            {
                var layer = doc.Layers[rhinoObject.Attributes.LayerIndex];
                if (!rhinoObject.Attributes.Visible || !layer.IsVisible) continue;
                var userStrings = rhinoObject.Attributes.GetUserStrings();
                if (userStrings == null || !userStrings.AllKeys.Contains("str-material"))
                {
                    idsToSelect.Add(rhinoObject.Id);
                }
            }

            doc.Views.RedrawEnabled = false;
            try
            {
                doc.Objects.UnselectAll();
                if (idsToSelect.Count > 0)
                {
                    doc.Objects.Select(idsToSelect, true);
                    RhinoApp.WriteLine($"{idsToSelect.Count} unassigned object(s) selected.");
                }
                else
                {
                    RhinoApp.WriteLine("No unassigned objects found.");
                }
            }
            finally
            {
                doc.Views.RedrawEnabled = true;
                doc.Views.Redraw();
            }
        }

        protected override void Dispose(bool disposing)
        {
            RhinoDoc.SelectObjects -= OnSelectionChanged;
            RhinoDoc.DeselectObjects -= OnSelectionChanged;
            RhinoDoc.DeselectAllObjects -= OnDeselectAllObjects;
            RhinoDoc.AddRhinoObject -= OnDatabaseChanged;
            RhinoDoc.DeleteRhinoObject -= OnDatabaseChanged;
            RhinoDoc.ReplaceRhinoObject -= OnReplaceObject;
            RhinoDoc.UserStringChanged -= OnUserStringChanged;
            RhinoDoc.EndOpenDocument -= OnDocumentChanged;
            RhinoDoc.NewDocument -= OnDocumentChanged;
            RhinoApp.Idle -= OnRhinoIdle;

            base.Dispose(disposing);
        }

        private void OnSelectionChanged(object sender, RhinoObjectSelectionEventArgs e) { if (!_isSyncingSelection) UpdatePanelData(); }
        private void OnDeselectAllObjects(object sender, RhinoDeselectAllObjectsEventArgs e) { if (!_isSyncingSelection) UpdatePanelData(); }
        private void OnDatabaseChanged(object sender, RhinoObjectEventArgs e) { if (!_isSyncingSelection) UpdatePanelData(); }
        private void OnReplaceObject(object sender, RhinoReplaceObjectEventArgs e) { if (!_isSyncingSelection) UpdatePanelData(); }
        private void OnUserStringChanged(object sender, RhinoDoc.UserStringChangedArgs e) { if (!_isSyncingSelection) UpdatePanelData(); }

        private void UpdatePanelData() => Eto.Forms.Application.Instance.AsyncInvoke(UpdateUserTextGrid);

        private void UpdateUserTextGrid()
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;

            IEnumerable<RhinoObject> objectsToProcess = _showAllObjectsCheckBox.Checked == true
                ? doc.Objects.Where(obj => obj.Attributes.Visible && doc.Layers[obj.Attributes.LayerIndex].IsVisible)
                : doc.Objects.GetSelectedObjects(false, false);

            var objectsList = objectsToProcess.ToList();
            if (objectsList.Count == 0)
            {
                _userTextGridView.DataStore = new List<UserTextEntry>();
                _totalLcaLabel.Text = "Total LCA: 0.00";
                _userTextGridView.Invalidate();
                return;
            }

            bool groupByMaterial = _groupByMaterialCheckBox.Checked == true;
            bool groupByClass = _groupByClassCheckBox.Checked == true;
            bool aggregate = _aggregateCheckBox.Checked == true;

            List<UserTextEntry> gridData;

            if (aggregate)
            {
                #region Aggregated View Logic
                var preProcessedObjects = objectsList.Select(o => {
                    TryComputeQuantity(o.Geometry, doc, out double qty, out string unit, out string qtyType);
                    return new
                    {
                        RhinoObject = o,
                        Material = o.Attributes.GetUserString("str-material") ?? "N/A",
                        IfcClass = o.Attributes.GetUserString("str-ifcclass") ?? "N/A",
                        Quantity = qty,
                        QuantityUnit = unit,
                        QuantityType = qtyType
                    };
                }).ToList();

                var aggregatedData = preProcessedObjects
                    .GroupBy(p => (p.Material, p.IfcClass, p.QuantityType))
                    .Select(g =>
                    {
                        double totalQuantity = g.Sum(p => p.Quantity);
                        double totalLca = 0.0;
                        double refQuantity = 0.0;
                        string refUnit = "";
                        double refLca = 0.0;

                        if (g.Key.QuantityType != "N/A" && _materialLcaData.TryGetValue(g.Key.Material, out MaterialData materialData))
                        {
                            totalLca = totalQuantity * materialData.Lca;
                            refQuantity = materialData.ReferenceQuantity;
                            refUnit = materialData.ReferenceUnit;
                            refLca = materialData.Lca;
                        }

                        return new UserTextEntry
                        {
                            Value = g.Key.Material,
                            IfcClass = g.Key.IfcClass,
                            ReferenceQuantity = refQuantity,
                            ReferenceUnit = refUnit,
                            ReferenceLca = refLca,
                            Count = g.Count(),
                            Quantity = totalQuantity,
                            QuantityUnit = g.First().QuantityUnit,
                            QuantityType = g.Key.QuantityType,
                            Lca = totalLca,
                            ObjectIds = g.Select(p => p.RhinoObject.Id).ToList()
                        };
                    });

                IEnumerable<UserTextEntry> sortedData;
                if (groupByClass)
                {
                    sortedData = aggregatedData.OrderBy(d => d.IfcClass == "N/A").ThenBy(d => d.IfcClass).ThenBy(d => d.Value);
                }
                else if (groupByMaterial)
                {
                    sortedData = aggregatedData.OrderBy(d => d.Value == "N/A").ThenBy(d => d.Value).ThenBy(d => d.IfcClass);
                }
                else
                {
                    sortedData = aggregatedData.OrderBy(d => d.Value == "N/A").ThenBy(d => d.Value);
                }

                gridData = FormatDataForHierarchy(sortedData.ToList(), groupByClass, groupByMaterial);
                #endregion
            }
            else
            {
                #region Non-Aggregated View Logic
                var processedObjects = objectsList.Select(obj =>
                {
                    TryComputeQuantity(obj.Geometry, doc, out double quantity, out string unit, out string quantityType);

                    string materialName = obj.Attributes.GetUserString("str-material") ?? "N/A";
                    double lca = 0.0;
                    double refQuantity = 0.0;
                    string refUnit = "";
                    double refLca = 0.0;

                    _quantityMultipliers.TryGetValue(obj.Id, out double multiplier);
                    if (multiplier < 1e-9) multiplier = 1.0;

                    if (_materialLcaData.TryGetValue(materialName, out MaterialData materialData))
                    {
                        refQuantity = materialData.ReferenceQuantity;
                        refUnit = materialData.ReferenceUnit;
                        refLca = materialData.Lca;

                        if (quantityType != "N/A")
                        {
                            lca = (quantity * multiplier) * materialData.Lca;
                        }
                    }

                    return new UserTextEntry
                    {
                        Value = materialName,
                        IfcClass = obj.Attributes.GetUserString("str-ifcclass") ?? "N/A",
                        ReferenceQuantity = refQuantity,
                        ReferenceUnit = refUnit,
                        ReferenceLca = refLca,
                        Count = 1,
                        Quantity = quantity,
                        QuantityUnit = unit,
                        QuantityType = quantityType,
                        QuantityMultiplier = multiplier,
                        Lca = lca,
                        ObjectIds = new List<Guid> { obj.Id }
                    };
                }).ToList();

                IEnumerable<UserTextEntry> sortedData;
                if (groupByClass)
                {
                    sortedData = processedObjects.OrderBy(d => d.IfcClass == "N/A").ThenBy(d => d.IfcClass).ThenBy(d => d.Value);
                }
                else if (groupByMaterial)
                {
                    sortedData = processedObjects.OrderBy(d => d.Value == "N/A").ThenBy(d => d.Value).ThenBy(d => d.IfcClass);
                }
                else
                {
                    sortedData = processedObjects.OrderBy(d => d.Value == "N/A").ThenBy(d => d.Value).ThenBy(d => d.IfcClass);
                }
                gridData = FormatDataForHierarchy(sortedData.ToList(), groupByClass, groupByMaterial);
                #endregion
            }

            var finalData = gridData;
            if (_showUnassignedCheckBox.Checked == false)
            {
                finalData.RemoveAll(entry => entry.Value == "N/A" && entry.IfcClass == "N/A");
            }

            _userTextGridView.DataStore = finalData;
            _totalLcaLabel.Text = $"Total LCA: {finalData.Sum(entry => entry.Lca):0.00} kgCO2 eq";
            _userTextGridView.Invalidate();

            try
            {
                _isSyncingSelection = true;
                var selectedObjectIds = doc.Objects.GetSelectedObjects(false, false).Select(o => o.Id).ToHashSet();
                if (!selectedObjectIds.Any())
                {
                    _userTextGridView.UnselectAll();
                    return;
                }
                var rowsToSelect = new List<int>();
                for (int i = 0; i < finalData.Count; i++)
                {
                    if (finalData[i].ObjectIds.Any(selectedObjectIds.Contains))
                    {
                        rowsToSelect.Add(i);
                    }
                }
                if (rowsToSelect.Any()) _userTextGridView.SelectedRows = rowsToSelect;
                else _userTextGridView.UnselectAll();
            }
            finally
            {
                _isSyncingSelection = false;
            }
        }

        private List<UserTextEntry> FormatDataForHierarchy(List<UserTextEntry> sortedData, bool groupByClass, bool groupByMaterial)
        {
            var formattedList = new List<UserTextEntry>();
            string lastPrimaryGroup = null;
            int sn = 1;

            foreach (var item in sortedData)
            {
                string currentPrimaryGroup = "N/A";
                if (groupByClass) currentPrimaryGroup = item.IfcClass;
                else if (groupByMaterial) currentPrimaryGroup = item.Value;
                else currentPrimaryGroup = sn.ToString();

                bool isFirstInGroup = currentPrimaryGroup != lastPrimaryGroup;
                item.SerialNumber = isFirstInGroup ? $"{sn++}." : "";

                if (groupByClass && !isFirstInGroup) item.IfcClass = "";
                if (groupByMaterial && !isFirstInGroup) item.Value = "";

                formattedList.Add(item);
                lastPrimaryGroup = currentPrimaryGroup;
            }
            return formattedList;
        }

        private string GetUnitType(string unit)
        {
            if (string.IsNullOrWhiteSpace(unit)) return "N/A";
            if (unit.EndsWith("3") || unit.ToLower().Contains("cu")) return "Volume";
            if (unit.EndsWith("2") || unit.ToLower().Contains("sq")) return "Area";
            return "Length";
        }

        private bool TryComputeQuantity(GeometryBase geo, RhinoDoc doc, out double quantity, out string unit, out string quantityType)
        {
            quantity = 0.0;
            unit = string.Empty;
            quantityType = "N/A";
            if (geo == null || doc == null) return false;

            var unitAbbreviation = doc.GetUnitSystemName(true, true, false, true);

            switch (geo)
            {
                case Brep brep when brep.IsSolid:
                    quantity = brep.GetVolume();
                    quantityType = "Volume";
                    unit = $"{unitAbbreviation}³";
                    break;
                case Extrusion extrusion:
                    var brepFromExtrusion = extrusion.ToBrep();
                    if (brepFromExtrusion != null && brepFromExtrusion.IsSolid)
                    {
                        quantity = brepFromExtrusion.GetVolume();
                        quantityType = "Volume";
                        unit = $"{unitAbbreviation}³";
                    }
                    break;
                case Mesh mesh when mesh.IsClosed:
                    quantity = mesh.Volume();
                    quantityType = "Volume";
                    unit = $"{unitAbbreviation}³";
                    break;
            }
            if (quantity > 1e-9) return true;

            AreaMassProperties amp = null;
            switch (geo)
            {
                case Brep brep: amp = AreaMassProperties.Compute(brep); break;
                case Surface srf: amp = AreaMassProperties.Compute(srf); break;
                case Mesh mesh: amp = AreaMassProperties.Compute(mesh); break;
                case Curve curve when curve.IsPlanar(): amp = AreaMassProperties.Compute(curve); break;
            }
            if (amp != null && amp.Area > 1e-9)
            {
                quantity = amp.Area;
                quantityType = "Area";
                unit = $"{unitAbbreviation}²";
                return true;
            }

            if (geo is Curve curveForLength)
            {
                quantity = curveForLength.GetLength();
                if (quantity > 1e-9)
                {
                    quantityType = "Length";
                    unit = unitAbbreviation;
                    return true;
                }
            }

            return false;
        }
        #endregion
    }
}