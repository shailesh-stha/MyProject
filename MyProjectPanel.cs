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
using System.Reflection;
using System.Runtime.InteropServices;
using MyProject.Properties;
using System.IO;

namespace MyProject
{
    [Guid("01D22D46-3C89-43A5-A92B-A1D6519A4244")]
    public class MyProjectPanel : Panel
    {
        // Private Fields
        private readonly List<string> _ifcClasses;
        private readonly Dictionary<string, MaterialData> _materialLcaData;
        private readonly Dictionary<Guid, double> _quantityMultipliers = new Dictionary<Guid, double>();
        private readonly GridView<UserTextEntry> _userTextGridView;
        private readonly GridColumn _qtyMultiplierColumn;
        private readonly Label _totalLcaLabel;
        private readonly ComboBox _ifcDropdown;
        private readonly ComboBox _materialDropdown;
        private readonly CheckBox _showAllObjectsCheckBox;
        private readonly CheckBox _showUnassignedCheckBox;
        private readonly CheckBox _groupByMaterialCheckBox;
        private readonly CheckBox _groupByClassCheckBox;
        private readonly CheckBox _aggregateCheckBox;
        private bool _isSyncingSelection = false;
        private bool _needsRefresh = false;

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

        public MyProjectPanel(uint documentSerialNumber)
        {
            _materialLcaData = CsvReader.ReadMaterialLcaDataFromResource("MyProject.Resources.materialListWithUnits.csv");
            _ifcClasses = new List<string> { "Wall", "Slab", "Beam", "Column", "Foundation", "Roof", "Stair", "Ramp", "Door", "Window", "Railing", "Covering" };
            Styles.Add<Label>("bold_label", label => label.Font = SystemFonts.Bold());

            _userTextGridView = new GridView<UserTextEntry>
            {
                ShowHeader = true,
                AllowMultipleSelection = true,
                Width = 550
            };

            _qtyMultiplierColumn = new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.QuantityMultiplier)) { TextAlignment = TextAlignment.Right }, HeaderText = "Qty. Multiplier", Editable = true, Width = 90 };

            _totalLcaLabel = new Label { Text = "Total LCA: 0.00", Style = "bold_label" };
            _ifcDropdown = new ComboBox { Width = 180, AutoComplete = true };
            _materialDropdown = new ComboBox { Width = 180, AutoComplete = true };
            _showAllObjectsCheckBox = new CheckBox { Text = "Show All/Selected", Checked = true };
            _showUnassignedCheckBox = new CheckBox { Text = "Show/Hide Unassigned", Checked = true };
            _groupByMaterialCheckBox = new CheckBox { Text = "Group by Material", Checked = false };
            _groupByClassCheckBox = new CheckBox { Text = "Group by Class", Checked = true };
            _aggregateCheckBox = new CheckBox { Text = "Aggregate", Checked = false };

            var mainLayout = new DynamicLayout { Spacing = new Size(6, 6), Padding = new Padding(10) };

            mainLayout.Add(CreateUpperLayout());
            mainLayout.Add(new Divider());
            mainLayout.Add(CreateIfcDefinitionLayout());
            mainLayout.Add(new Divider());
            mainLayout.Add(CreateMaterialDefinitionLayout());
            mainLayout.Add(new Divider());
            mainLayout.Add(CreateDataLayout());
            mainLayout.Add(new Divider());
            mainLayout.Add(null, true);

            Content = mainLayout;
            MinimumSize = new Size(400, 550);

            RegisterEventHandlers();
            UpdatePanelData();
        }

        #region UI Layout Methods
        private Control CreateUpperLayout()
        {
            var structureIcon = BytesToEtoBitmap(Resources.strLCA256, new Size(18, 18));
            var btnStructure = new Button
            {
                Image = structureIcon,
                //Text = "structure.com", // or string.Empty,
                ToolTip = "Visit str-ucture.com",
                MinimumSize = Size.Empty,
                ImagePosition = ButtonImagePosition.Left
            };

            btnStructure.Click += (s, e) =>
            {
                var url = "http://www.str-ucture.com";
                try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
                catch (Exception ex) { RhinoApp.WriteLine($"Error opening website: {ex.Message}"); }
            };

            var buttonLayout = new StackLayout
            {
                Orientation = Orientation.Horizontal,
                Spacing = 5,
                Items = { null, btnStructure }
            };

            var layout = new DynamicLayout { Spacing = new Size(6, 6) };
            layout.AddRow(new Label { Text = "Utility Buttons", Style = "bold_label" });
            layout.AddRow(buttonLayout);
            return layout;
        }

        private Control CreateIfcDefinitionLayout()
        {
            _ifcDropdown.Items.Add(new ListItem { Text = "" });
            _ifcClasses.ForEach(cls => _ifcDropdown.Items.Add(new ListItem { Text = cls }));

            var assignIfcClassIcon = BytesToEtoBitmap(Resources.assignIfcClass256, new Size(18, 18));
            var assignButton= new Button
            {
                Image = assignIfcClassIcon,
                MinimumSize = Size.Empty,
                ImagePosition = ButtonImagePosition.Left
            };
            assignButton.Click += (s, e) => AssignUserString("str-ifcclass", _ifcDropdown.SelectedValue as ListItem);

            var removeIfcClassIcon = BytesToEtoBitmap(Resources.removeIfcClass256, new Size(18, 18));
            var removeButton = new Button
            {
                Image = removeIfcClassIcon,
                MinimumSize = Size.Empty,
                ImagePosition = ButtonImagePosition.Left
            };
            removeButton.Click += (s, e) => RemoveUserString("str-ifcclass");

            var buttonsLayout = new StackLayout
            {
                Orientation = Orientation.Horizontal,
                Spacing = 5,
                Items = { assignButton, removeButton }
            };

            var layout = new DynamicLayout { Spacing = new Size(6, 6) };
            layout.AddRow(new Label { Text = "IFC Class Definition", Style = "bold_label" });
            layout.AddRow(new Label { Text = "Select IfcClass:" }, _ifcDropdown, buttonsLayout);
            return layout;
        }

        private Control CreateMaterialDefinitionLayout()
        {
            _materialDropdown.Items.Add(new ListItem { Text = "" });
            _materialLcaData.Keys.ToList().ForEach(m => _materialDropdown.Items.Add(new ListItem { Text = m }));


            var assignMaterialIcon = BytesToEtoBitmap(Resources.assignMaterial256, new Size(16, 16));
            var assignButton = new Button
            {
                Image = assignMaterialIcon,
                MinimumSize = Size.Empty,
                ImagePosition = ButtonImagePosition.Left
            };
            assignButton.Click += (s, e) => AssignUserString("str-material", _materialDropdown.SelectedValue as ListItem);

            var removeMaterialIcon = BytesToEtoBitmap(Resources.removeMaterial256, new Size(16, 16));
            var removeButton = new Button
            {
                Image = removeMaterialIcon,
                MinimumSize = Size.Empty,
                ImagePosition = ButtonImagePosition.Left
            };
            removeButton.Click += (s, e) => RemoveUserString("str-material");

            var buttonsLayout = new StackLayout
            {
                Orientation = Orientation.Horizontal,
                Spacing = 5,
                Items = { assignButton, removeButton }
            };

            var layout = new DynamicLayout { Spacing = new Size(6, 6) };
            layout.AddRow(new Label { Text = "Material Definition", Style = "bold_label" });
            layout.AddRow(new Label { Text = "Select Material:" }, _materialDropdown, buttonsLayout);
            return layout;
        }

        private Control CreateDataLayout()
        {
            _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.SerialNumber)), HeaderText = "SN", Editable = false, Width = 30 });
            _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.IfcClass)), HeaderText = "IfcClass", Editable = false, Width = 60 });
            _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.Value)), HeaderText = "Material (Name)", Editable = false, Width = 150 });
            _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.DisplayReference)) { TextAlignment = TextAlignment.Right }, HeaderText = "Ref. Qty.", Editable = false, Width = 70 });
            _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell { Binding = Binding.Property<UserTextEntry, string>(entry => $"{entry.ReferenceLca:0.00}"), TextAlignment = TextAlignment.Right }, HeaderText = "Ref. LCA", Editable = false, Width = 70 });
            _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.Count)) { TextAlignment = TextAlignment.Right }, HeaderText = "Count", Editable = false, Width = 50 });
            _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.DisplayQuantity)) { TextAlignment = TextAlignment.Right }, HeaderText = "Qty. (Rh)", Editable = false, Width = 75 });
            _userTextGridView.Columns.Add(_qtyMultiplierColumn);
            _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell { Binding = Binding.Property<UserTextEntry, string>(entry => $"{entry.QuantityTotal:0.00}"), TextAlignment = TextAlignment.Right }, HeaderText = "Qty. Total", Editable = false, Width = 75 });
            _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell { Binding = Binding.Property<UserTextEntry, string>(entry => $"{entry.Lca:0.00}"), TextAlignment = TextAlignment.Right }, HeaderText = "LCA (kgCO2 eq)", Editable = false, Width = 100 });

            _qtyMultiplierColumn.Visible = !_aggregateCheckBox.Checked ?? true;

            _showAllObjectsCheckBox.CheckedChanged += (s, e) => UpdatePanelData();
            _showUnassignedCheckBox.CheckedChanged += (s, e) => UpdatePanelData();
            _aggregateCheckBox.CheckedChanged += OnAggregateChanged;
            _groupByMaterialCheckBox.CheckedChanged += OnGroupByMaterialChanged;
            _groupByClassCheckBox.CheckedChanged += OnGroupByClassChanged;

            _userTextGridView.SelectionChanged += OnGridSelectionChanged;
            _userTextGridView.CellEdited += OnGridCellEdited;

            var scrollableGrid = new Scrollable { Content = _userTextGridView, Height = 250 };
            var viewOptionsLayout = new StackLayout { Orientation = Orientation.Horizontal, Spacing = 10, Items = { _showAllObjectsCheckBox, _showUnassignedCheckBox } };
            var groupingOptionsLayout = new StackLayout { Orientation = Orientation.Horizontal, Spacing = 10, Items = { _groupByClassCheckBox, _groupByMaterialCheckBox, _aggregateCheckBox } };

            var btnTableColumns = new Button { Text = "Table Columns" };
            btnTableColumns.Click += (s, e) => new ColumnVisibilityDialog(_userTextGridView).ShowModal(this);

            var btnSelectUnassigned = new Button { Text = "Select Unassigned" };
            btnSelectUnassigned.Click += OnSelectUnassignedClick;

            var gridHeaderLayout = new DynamicLayout { Spacing = new Size(6, 6) };
            gridHeaderLayout.AddRow(new Label { Text = "Attribute User Text", Style = "bold_label" }, null, btnSelectUnassigned, btnTableColumns);

            var layout = new DynamicLayout { Spacing = new Size(6, 6) };
            layout.AddRow(gridHeaderLayout);
            layout.AddRow(viewOptionsLayout);
            layout.AddRow(groupingOptionsLayout);
            layout.AddRow(scrollableGrid);
            layout.AddRow(_totalLcaLabel, null);
            return layout;
        }

        /// <summary>
        /// **MODIFIED**: Helper method to convert a byte array from project resources into an Eto.Drawing.Bitmap, with optional resizing.
        /// </summary>
        /// <param name="bytes">The byte array containing the icon data.</param>
        /// <param name="desiredSize">An optional Eto.Drawing.Size to resize the icon to. If null, the original size is used.</param>
        /// <returns>An Eto.Drawing.Bitmap object if conversion is successful, otherwise null.</returns>
        private Bitmap BytesToEtoBitmap(byte[] bytes, Size? desiredSize = null)
        {
            try
            {
                if (bytes == null || bytes.Length == 0)
                {
                    RhinoApp.WriteLine("Error: Icon resource byte array is null or empty.");
                    return null;
                }
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

        #region Event Handlers
        private void RegisterEventHandlers()
        {
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

        private void OnAggregateChanged(object sender, EventArgs e)
        {
            _qtyMultiplierColumn.Visible = !_aggregateCheckBox.Checked ?? true;
            UpdatePanelData();
        }

        private void OnGroupByClassChanged(object sender, EventArgs e)
        {
            if (_groupByClassCheckBox.Checked == true)
            {
                _groupByMaterialCheckBox.Checked = false;
            }
            UpdatePanelData();
        }

        private void OnGroupByMaterialChanged(object sender, EventArgs e)
        {
            if (_groupByMaterialCheckBox.Checked == true)
            {
                _groupByClassCheckBox.Checked = false;
            }
            ReorderGridColumns(_groupByMaterialCheckBox.Checked ?? false);
            UpdatePanelData();
        }

        private void OnSelectUnassignedClick(object sender, EventArgs e)
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;

            var idsToSelect = doc.Objects.Where(o => o.IsSelectable(true, false, false, true) &&
                                                     o.Attributes.Visible &&
                                                     doc.Layers[o.Attributes.LayerIndex].IsVisible &&
                                                     string.IsNullOrWhiteSpace(o.Attributes.GetUserString("str-material")))
                                        .Select(o => o.Id)
                                        .ToList();

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

        private void OnGridCellEdited(object sender, GridViewCellEventArgs e)
        {
            if (e.Item is UserTextEntry entry && entry.ObjectIds.Any())
            {
                var objectId = entry.ObjectIds.First();
                _quantityMultipliers[objectId] = entry.QuantityMultiplier;
                UpdatePanelData();
            }
        }

        private void OnGridSelectionChanged(object sender, EventArgs e)
        {
            if (_isSyncingSelection) return;
            _isSyncingSelection = true;

            var doc = RhinoDoc.ActiveDoc;
            if (doc == null || _userTextGridView.DataStore == null)
            {
                _isSyncingSelection = false;
                return;
            }

            var idsToSelect = _userTextGridView.SelectedItems.SelectMany(entry => entry.ObjectIds).ToHashSet();
            doc.Objects.UnselectAll();
            doc.Objects.Select(idsToSelect, true);
            doc.Views.Redraw();

            _isSyncingSelection = false;
        }

        private void OnSelectionChanged(object sender, RhinoObjectSelectionEventArgs e) => UpdatePanelDataSafe();
        private void OnDeselectAllObjects(object sender, RhinoDeselectAllObjectsEventArgs e) => UpdatePanelDataSafe();
        private void OnDatabaseChanged(object sender, RhinoObjectEventArgs e) => UpdatePanelDataSafe();
        private void OnReplaceObject(object sender, RhinoReplaceObjectEventArgs e) => UpdatePanelDataSafe();
        private void OnUserStringChanged(object sender, RhinoDoc.UserStringChangedArgs e) => UpdatePanelDataSafe();

        private void UpdatePanelDataSafe()
        {
            if (!_isSyncingSelection) UpdatePanelData();
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

        #endregion

        #region Core Logic Methods
        private void UpdatePanelData() => Eto.Forms.Application.Instance.AsyncInvoke(UpdateUserTextGrid);

        private void UpdateUserTextGrid()
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null)
            {
                _userTextGridView.DataStore = new List<UserTextEntry>();
                _totalLcaLabel.Text = "Total LCA: 0.00";
                _userTextGridView.Invalidate();
                return;
            }

            var objectsToProcess = (_showAllObjectsCheckBox.Checked == true
                ? doc.Objects.Where(obj => obj.Attributes.Visible && doc.Layers[obj.Attributes.LayerIndex].IsVisible)
                : doc.Objects.GetSelectedObjects(false, false)).ToList();

            if (!objectsToProcess.Any())
            {
                _userTextGridView.DataStore = new List<UserTextEntry>();
                _totalLcaLabel.Text = "Total LCA: 0.00";
                _userTextGridView.Invalidate();
                return;
            }

            var processedObjects = objectsToProcess.Select(o =>
            {
                TryComputeQuantity(o.Geometry, doc, out double qty, out string unit, out string qtyType);
                return new
                {
                    RhinoObject = o,
                    Material = o.Attributes.GetUserString("str-material") ?? "N/A",
                    IfcClass = o.Attributes.GetUserString("str-ifcclass") ?? "N/A",
                    Quantity = qty,
                    QuantityUnit = unit,
                    QuantityType = qtyType,
                    QuantityMultiplier = _quantityMultipliers.ContainsKey(o.Id) ? _quantityMultipliers[o.Id] : 1.0
                };
            }).ToList();

            List<UserTextEntry> gridData;

            if (_aggregateCheckBox.Checked == true)
            {
                var aggregatedData = processedObjects
                    .GroupBy(p => new { p.Material, p.IfcClass, p.QuantityType })
                    .Select(g =>
                    {
                        var first = g.First();
                        var totalQuantity = g.Sum(p => p.Quantity);
                        var totalLca = 0.0;
                        var refQuantity = 0.0;
                        var refUnit = "";
                        var refLca = 0.0;

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
                            QuantityUnit = first.QuantityUnit,
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
                    double refQuantity = 0.0;
                    string refUnit = "";
                    double refLca = 0.0;

                    if (p.QuantityType != "N/A" && _materialLcaData.TryGetValue(p.Material, out MaterialData materialData))
                    {
                        lca = (p.Quantity * p.QuantityMultiplier) * materialData.Lca;
                        refQuantity = materialData.ReferenceQuantity;
                        refUnit = materialData.ReferenceUnit;
                        refLca = materialData.Lca;
                    }

                    return new UserTextEntry
                    {
                        Value = p.Material,
                        IfcClass = p.IfcClass,
                        ReferenceQuantity = refQuantity,
                        ReferenceUnit = refUnit,
                        ReferenceLca = refLca,
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
                gridData.RemoveAll(entry => entry.Value == "N/A" && entry.IfcClass == "N/A");
            }

            _userTextGridView.DataStore = gridData;
            _totalLcaLabel.Text = $"Total LCA: {gridData.Sum(entry => entry.Lca):0.00} kgCO2 eq";
            _userTextGridView.Invalidate();

            SyncGridSelectionWithRhino();
        }

        private List<UserTextEntry> SortAndFormatData(List<UserTextEntry> data)
        {
            IEnumerable<UserTextEntry> sortedData;
            if (_groupByClassCheckBox.Checked == true)
            {
                sortedData = data.OrderBy(d => d.IfcClass == "N/A").ThenBy(d => d.IfcClass).ThenBy(d => d.Value);
            }
            else if (_groupByMaterialCheckBox.Checked == true)
            {
                sortedData = data.OrderBy(d => d.Value == "N/A").ThenBy(d => d.Value).ThenBy(d => d.IfcClass);
            }
            else
            {
                sortedData = data.OrderBy(d => d.Value == "N/A").ThenBy(d => d.Value);
            }

            var formattedList = new List<UserTextEntry>();
            string lastPrimaryGroup = null;
            int serialNumber = 1;

            foreach (var item in sortedData)
            {
                string currentPrimaryGroup = GetCurrentGroupingKey(item);
                bool isFirstInGroup = currentPrimaryGroup != lastPrimaryGroup;

                item.SerialNumber = isFirstInGroup ? $"{serialNumber++}." : "";

                if (isFirstInGroup)
                {
                    if (_groupByClassCheckBox.Checked == true) item.Value = item.Value;
                    else if (_groupByMaterialCheckBox.Checked == true) item.IfcClass = item.IfcClass;
                }
                else
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
            _isSyncingSelection = true;
            try
            {
                var doc = RhinoDoc.ActiveDoc;
                if (doc == null) return;

                var selectedObjectIds = doc.Objects.GetSelectedObjects(false, false).Select(o => o.Id).ToHashSet();
                var rowsToSelect = new List<int>();

                if (_userTextGridView.DataStore != null)
                {
                    var dataStoreList = _userTextGridView.DataStore.ToList();
                    for (int i = 0; i < dataStoreList.Count; i++)
                    {
                        if (dataStoreList[i].ObjectIds.Any(selectedObjectIds.Contains))
                        {
                            rowsToSelect.Add(i);
                        }
                    }
                }

                _userTextGridView.SelectedRows = rowsToSelect;
            }
            finally
            {
                _isSyncingSelection = false;
            }
        }

        private void AssignUserString(string key, ListItem selectedItem)
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;
            var value = selectedItem?.Text;

            if (string.IsNullOrWhiteSpace(value))
            {
                RhinoApp.WriteLine($"Please select a valid {key.Substring(4)} from the dropdown first.");
                return;
            }

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
                newAttributes.SetUserString(key, value);
                if (doc.Objects.ModifyAttributes(rhinoObject.Id, newAttributes, false)) updatedCount++;
            }

            doc.Views.Redraw();
            RhinoApp.WriteLine($"Successfully assigned '{value}' to {updatedCount} object(s).");
        }

        private void RemoveUserString(string key)
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;

            var selectedObjects = doc.Objects.GetSelectedObjects(false, false).ToList();
            if (!selectedObjects.Any())
            {
                RhinoApp.WriteLine($"Please select one or more objects to remove the '{key.Substring(4)}' definition from.");
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
                doc.Views.Redraw();
                RhinoApp.WriteLine($"Successfully removed '{key.Substring(4)}' definition from {updatedCount} object(s).");
            }
            else
            {
                RhinoApp.WriteLine($"Selected objects do not have a '{key.Substring(4)}' definition to remove.");
            }
        }

        private void ReorderGridColumns(bool groupByMaterial)
        {
            var ifcClassColumn = _userTextGridView.Columns.FirstOrDefault(c => c.HeaderText == "IfcClass");
            var materialNameColumn = _userTextGridView.Columns.FirstOrDefault(c => c.HeaderText == "Material (Name)");

            if (ifcClassColumn == null || materialNameColumn == null) return;

            int ifcIndex = _userTextGridView.Columns.IndexOf(ifcClassColumn);
            int materialIndex = _userTextGridView.Columns.IndexOf(materialNameColumn);

            if (groupByMaterial && ifcIndex < materialIndex)
            {
                _userTextGridView.Columns.RemoveAt(ifcIndex);
                _userTextGridView.Columns.Insert(materialIndex, ifcClassColumn);
            }
            else if (!groupByMaterial && ifcIndex > materialIndex)
            {
                _userTextGridView.Columns.RemoveAt(ifcIndex);
                _userTextGridView.Columns.Insert(materialIndex, ifcClassColumn);
            }
        }

        private bool TryComputeQuantity(GeometryBase geo, RhinoDoc doc, out double quantity, out string unit, out string quantityType)
        {
            quantity = 0.0;
            unit = string.Empty;
            quantityType = "N/A";
            if (geo == null || doc == null) return false;

            var unitAbbreviation = doc.GetUnitSystemName(true, true, false, true);

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

        private bool TryGetVolume(GeometryBase geo, out double quantity)
        {
            quantity = 0.0;
            switch (geo)
            {
                case Brep brep when brep.IsSolid:
                    quantity = brep.GetVolume();
                    break;
                case Extrusion extrusion:
                    var brepFromExtrusion = extrusion.ToBrep();
                    if (brepFromExtrusion != null && brepFromExtrusion.IsSolid) quantity = brepFromExtrusion.GetVolume();
                    break;
                case Mesh mesh when mesh.IsClosed:
                    quantity = mesh.Volume();
                    break;
            }
            return quantity > 1e-9;
        }

        private bool TryGetArea(GeometryBase geo, out double quantity)
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

        private bool TryGetLength(GeometryBase geo, out double quantity)
        {
            quantity = 0.0;
            if (geo is Curve curveForLength) quantity = curveForLength.GetLength();
            return quantity > 1e-9;
        }

        #endregion
    }
}