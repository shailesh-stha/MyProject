// MyProjectPanel.cs
using Eto.Drawing;
using Eto.Forms;
using MyProject.Properties;
using Rhino;
using Rhino.DocObjects;
using Rhino.Geometry;
using Rhino.UI.Controls;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace MyProject
{
    /// <summary>
    /// Represents the main Eto.Forms panel for the MyProject plugin.
    /// This panel displays object attributes, LCA data, and document information.
    /// </summary>
    [Guid("01D22D46-3C89-43A5-A92B-A1D6519A4244")]
    public class MyProjectPanel : Panel
    {
        #region Private Constants
        private const string MaterialKey = "STR_MATERIAL";
        private const string IfcClassKey = "STR_IFC_CLASS";
        #endregion

        #region Private Fields
        private List<string> _ifcClasses = new List<string>();
        private readonly Dictionary<string, MaterialData> _materialLcaData;
        private readonly Dictionary<Guid, double> _quantityMultipliers = new Dictionary<Guid, double>();

        // UI Controls are initialized here instead of in the constructor for clarity.
        private readonly GridView<UserTextEntry> _userTextGridView = new GridView<UserTextEntry> { ShowHeader = true, AllowMultipleSelection = true, Width = 550 };
        private readonly GridView<DocumentUserTextEntry> _docUserTextGridView = new GridView<DocumentUserTextEntry> { ShowHeader = true, Height = 100, Width = 550 };
        private readonly GridColumn _qtyMultiplierColumn = new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.QuantityMultiplier)) { TextAlignment = TextAlignment.Right }, HeaderText = "Qty. Multiplier", Editable = true, Width = 90 };
        private readonly Label _totalLcaLabel = new Label { Text = "Total LCA: 0.00" };
        private readonly ComboBox _ifcDropdown = new ComboBox { Width = 180, AutoComplete = true };
        private readonly ComboBox _materialDropdown = new ComboBox { Width = 180, AutoComplete = true };
        private readonly CheckBox _showAllObjectsCheckBox = new CheckBox { Text = "Show All/Selected", Checked = true };
        private readonly CheckBox _showUnassignedCheckBox = new CheckBox { Text = "Show/Hide Unassigned", Checked = true };
        private readonly CheckBox _groupByMaterialCheckBox = new CheckBox { Text = "Group by Material", Checked = false };
        private readonly CheckBox _groupByClassCheckBox = new CheckBox { Text = "Group by Class", Checked = true };
        private readonly CheckBox _aggregateCheckBox = new CheckBox { Text = "Aggregate", Checked = false };
        private TextArea _notesTextArea;

        private bool _isSyncingSelection = false;
        private bool _needsRefresh = false;
        private bool _isMouseOverGrid = false;
        #endregion

        #region Data-Binding Classes
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
        #endregion

        /// <summary>
        /// Constructor for the panel. Required by the Rhino panel framework.
        /// </summary>
#pragma warning disable IDE0060 // Remove unused parameter
        public MyProjectPanel(uint documentSerialNumber)
#pragma warning restore IDE0060 // Remove unused parameter
        {
            _materialLcaData = CsvReader.ReadMaterialLcaDataFromResource("MyProject.Resources.Data.materialListWithUnits.csv");

            InitializeLayout();
            RegisterEventHandlers();

            ReloadIfcClassList();
            UpdatePanelData();
        }

        /// <summary>
        /// Sets up the entire UI layout for the panel.
        /// </summary>
        private void InitializeLayout()
        {
            Styles.Add<Label>("bold_label", label =>
            {
                label.Font = SystemFonts.Bold();
                _totalLcaLabel.Font = label.Font; // Apply style to total label as well
            });

            var mainLayout = new DynamicLayout { Spacing = new Size(6, 6), Padding = new Padding(10) };

            mainLayout.Add(new Expander
            {
                Header = new Label { Text = "Utility Buttons", Style = "bold_label" },
                Content = CreateUtilityButtonsLayout(),
                Expanded = false
            });

            mainLayout.Add(new Expander
            {
                Header = new Label { Text = "Definition", Style = "bold_label" },
                Content = CreateDefinitionLayout(),
                Expanded = true
            });

            mainLayout.Add(new Expander
            {
                Header = new Label { Text = "Attribute User Text", Style = "bold_label" },
                Content = CreateAttributeGridLayout(),
                Expanded = true
            });

            mainLayout.Add(new Expander
            {
                Header = new Label { Text = "Document Notes/User Text", Style = "bold_label" },
                Content = CreateCombinedNotesAndUserTextLayout(),
                Expanded = false
            });

            mainLayout.Add(null, true);

            Content = new Scrollable { Content = mainLayout, Border = BorderType.None };
            MinimumSize = new Size(400, 550);
        }

        #region UI Layout Methods
        private Control CreateUtilityButtonsLayout()
        {
            var structureIcon = BytesToEtoBitmap(Resources.btn_strLCA256, new Size(18, 18));
            var btnStructure = new Button
            {
                Image = structureIcon,
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
            layout.AddRow(buttonLayout);
            return layout;
        }

        private Control CreateCombinedNotesAndUserTextLayout()
        {
            var layout = new DynamicLayout { Spacing = new Size(6, 6) };

            // --- Document Notes Part ---
            _notesTextArea = new TextArea
            {
                Width = 550,
                Height = 80,
                Text = RhinoDoc.ActiveDoc?.Notes ?? string.Empty
            };

            var saveButton = new Button { Text = "Save Notes" };
            saveButton.Click += (s, e) =>
            {
                var doc = RhinoDoc.ActiveDoc;
                if (doc != null)
                {
                    doc.Notes = _notesTextArea.Text;
                    RhinoApp.WriteLine("Document notes saved.");
                }
            };

            layout.AddRow(_notesTextArea);
            layout.AddRow(saveButton, null);
            layout.Add(new Divider());

            // --- Document User Text Part ---
            _docUserTextGridView.Columns.Add(new GridColumn { HeaderText = "Key", DataCell = new TextBoxCell(nameof(DocumentUserTextEntry.Key)), Editable = false, Width = 150 });
            _docUserTextGridView.Columns.Add(new GridColumn { HeaderText = "Value", DataCell = new TextBoxCell(nameof(DocumentUserTextEntry.Value)), Editable = false });
            var refreshButton = new Button { Text = "Refresh User Text" };
            refreshButton.Click += (s, e) => UpdatePanelData();
            layout.AddRow(new Scrollable { Content = _docUserTextGridView });
            layout.AddRow(refreshButton);

            return layout;
        }

        private Control CreateDefinitionLayout()
        {
            var layout = new DynamicLayout { Spacing = new Size(6, 6) };

            // --- IFC Class Part ---
            var assignIfcButton = new Button { Image = BytesToEtoBitmap(Resources.btn_assignIfcClass256, new Size(18, 18)), MinimumSize = Size.Empty, ImagePosition = ButtonImagePosition.Left };
            assignIfcButton.Click += (s, e) => AssignUserString(IfcClassKey, "IfcClass", _ifcDropdown.SelectedValue as ListItem);
            var removeIfcButton = new Button { Image = BytesToEtoBitmap(Resources.btn_removeIfcClass256, new Size(18, 18)), MinimumSize = Size.Empty, ImagePosition = ButtonImagePosition.Left };
            removeIfcButton.Click += (s, e) => RemoveUserString(IfcClassKey, "IfcClass");
            var refreshIfcButton = new Button { Text = "Refresh" };
            refreshIfcButton.Click += (s, e) => ReloadIfcClassList();
            var ifcButtonsLayout = new StackLayout { Orientation = Orientation.Horizontal, Spacing = 5, Items = { assignIfcButton, removeIfcButton, refreshIfcButton } };
            layout.AddRow(new Label { Text = "Select IfcClass:" }, _ifcDropdown, ifcButtonsLayout);

            // --- Material Part ---
            // MODIFIED: This block of code was missing and has been restored.
            _materialDropdown.Items.Clear();
            _materialDropdown.Items.Add(new ListItem { Text = "" });
            foreach (var materialName in _materialLcaData.Keys)
            {
                _materialDropdown.Items.Add(new ListItem { Text = materialName });
            }

            var assignMaterialButton = new Button { Image = BytesToEtoBitmap(Resources.btn_assignMaterial256, new Size(16, 16)), MinimumSize = Size.Empty, ImagePosition = ButtonImagePosition.Left };
            assignMaterialButton.Click += (s, e) => AssignUserString(MaterialKey, "Material", _materialDropdown.SelectedValue as ListItem);
            var removeMaterialButton = new Button { Image = BytesToEtoBitmap(Resources.btn_removeMaterial256, new Size(16, 16)), MinimumSize = Size.Empty, ImagePosition = ButtonImagePosition.Left };
            removeMaterialButton.Click += (s, e) => RemoveUserString(MaterialKey, "Material");
            var materialButtonsLayout = new StackLayout { Orientation = Orientation.Horizontal, Spacing = 5, Items = { assignMaterialButton, removeMaterialButton } };
            layout.AddRow(new Label { Text = "Select Material:" }, _materialDropdown, materialButtonsLayout);

            return layout;
        }

        private Control CreateAttributeGridLayout()
        {
            if (_userTextGridView.Columns.Count == 0)
            {
                _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.SerialNumber)), HeaderText = "SN", Editable = false, Width = 30 });
                _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.IfcClass)), HeaderText = "IfcClass", Editable = false, Width = 65 });
                _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.Value)), HeaderText = "Material (Name)", Editable = false, Width = 150 });
                _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.DisplayReference)) { TextAlignment = TextAlignment.Right }, HeaderText = "Ref. Qty.", Editable = false, Width = 70 });
                _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell { Binding = Eto.Forms.Binding.Property<UserTextEntry, string>(entry => $"{entry.ReferenceLca:0.00}"), TextAlignment = TextAlignment.Right }, HeaderText = "Ref. LCA", Editable = false, Width = 70 });
                _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.Count)) { TextAlignment = TextAlignment.Right }, HeaderText = "Count", Editable = false, Width = 50, Visible = false });
                _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.DisplayQuantity)) { TextAlignment = TextAlignment.Right }, HeaderText = "Qty. (Rh)", Editable = false, Width = 75 });
                _userTextGridView.Columns.Add(_qtyMultiplierColumn);
                _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell { Binding = Eto.Forms.Binding.Property<UserTextEntry, string>(entry => $"{entry.QuantityTotal:0.00}"), TextAlignment = TextAlignment.Right }, HeaderText = "Qty. Total", Editable = false, Width = 75 });
                _userTextGridView.Columns.Add(new GridColumn { DataCell = new TextBoxCell { Binding = Eto.Forms.Binding.Property<UserTextEntry, string>(entry => $"{entry.Lca:0.00}"), TextAlignment = TextAlignment.Right }, HeaderText = "LCA (kgCO2 eq)", Editable = false, Width = 100 });
            }

            _qtyMultiplierColumn.Visible = !_aggregateCheckBox.Checked ?? true;

            _showAllObjectsCheckBox.CheckedChanged += (s, e) => UpdatePanelData();
            _showUnassignedCheckBox.CheckedChanged += (s, e) => UpdatePanelData();
            _aggregateCheckBox.CheckedChanged += OnAggregateChanged;
            _groupByMaterialCheckBox.CheckedChanged += OnGroupByMaterialChanged;
            _groupByClassCheckBox.CheckedChanged += OnGroupByClassChanged;
            _userTextGridView.SelectionChanged += OnGridSelectionChanged;
            _userTextGridView.CellEdited += OnGridCellEdited;
            _userTextGridView.MouseEnter += (s, e) => _isMouseOverGrid = true;
            _userTextGridView.MouseLeave += (s, e) => _isMouseOverGrid = false;

            var columnsViewButton = new Button { Image = BytesToEtoBitmap(Resources.btn_columnsView256, new Size(16, 16)), MinimumSize = Size.Empty, ImagePosition = ButtonImagePosition.Left };
            columnsViewButton.Click += (s, e) => new ColumnVisibilityDialog(_userTextGridView).ShowModal(this);
            var selectUnassignedButton = new Button { Text = "Select Unassigned" };
            selectUnassignedButton.Click += OnSelectUnassignedClick;

            var gridHeaderLayout = new DynamicLayout { Spacing = new Size(6, 6) };
            gridHeaderLayout.AddRow(null, selectUnassignedButton, columnsViewButton);

            var viewOptionsLayout = new StackLayout { Orientation = Orientation.Horizontal, Spacing = 10, Items = { _showAllObjectsCheckBox, _showUnassignedCheckBox } };
            var groupingOptionsLayout = new StackLayout { Orientation = Orientation.Horizontal, Spacing = 10, Items = { _groupByClassCheckBox, _groupByMaterialCheckBox, _aggregateCheckBox } };

            var layout = new DynamicLayout { Spacing = new Size(6, 6) };
            layout.AddRow(gridHeaderLayout);
            layout.AddRow(viewOptionsLayout);
            layout.AddRow(groupingOptionsLayout);
            layout.AddRow(new Scrollable { Content = _userTextGridView, Height = 250 });
            layout.AddRow(_totalLcaLabel, null);
            return layout;
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

        #region Event Handlers
        private void RegisterEventHandlers()
        {
            RhinoDoc.SelectObjects += OnDocumentStateChanged;
            RhinoDoc.DeselectObjects += OnDocumentStateChanged;
            RhinoDoc.DeselectAllObjects += OnDocumentStateChanged;
            RhinoDoc.AddRhinoObject += OnDocumentStateChanged;
            RhinoDoc.DeleteRhinoObject += OnDocumentStateChanged;
            RhinoDoc.ReplaceRhinoObject += OnDocumentStateChanged;
            RhinoDoc.UserStringChanged += OnDocumentStateChanged;
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
            if (_notesTextArea != null)
            {
                _notesTextArea.Text = RhinoDoc.ActiveDoc?.Notes ?? string.Empty;
            }
            ReloadIfcClassList();
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
                _quantityMultipliers[entry.ObjectIds.First()] = entry.QuantityMultiplier;
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

        private void OnDocumentStateChanged(object sender, EventArgs e) => UpdatePanelDataSafe();

        private void UpdatePanelDataSafe()
        {
            if (!_isSyncingSelection) UpdatePanelData();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                RhinoDoc.SelectObjects -= OnDocumentStateChanged;
                RhinoDoc.DeselectObjects -= OnDocumentStateChanged;
                RhinoDoc.DeselectAllObjects -= OnDocumentStateChanged;
                RhinoDoc.AddRhinoObject -= OnDocumentStateChanged;
                RhinoDoc.DeleteRhinoObject -= OnDocumentStateChanged;
                RhinoDoc.ReplaceRhinoObject -= OnDocumentStateChanged;
                RhinoDoc.UserStringChanged -= OnDocumentStateChanged;
                RhinoDoc.EndOpenDocument -= OnDocumentChanged;
                RhinoDoc.NewDocument -= OnDocumentChanged;
                RhinoApp.Idle -= OnRhinoIdle;
            }
            base.Dispose(disposing);
        }

        #endregion

        #region Core Logic Methods
        private void UpdatePanelData() => Application.Instance.AsyncInvoke(UpdateUserTextGrid);

        private void UpdateUserTextGrid()
        {
            var doc = RhinoDoc.ActiveDoc;
            UpdateDocumentUserTextGrid(doc);

            if (doc == null)
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
                _quantityMultipliers.TryGetValue(o.Id, out double multiplier);
                return new
                {
                    RhinoObject = o,
                    Material = o.Attributes.GetUserString(MaterialKey) ?? "N/A",
                    IfcClass = o.Attributes.GetUserString(IfcClassKey) ?? "N/A",
                    Quantity = qty,
                    QuantityUnit = unit,
                    QuantityType = qtyType,
                    QuantityMultiplier = multiplier == 0 ? 1.0 : multiplier
                };
            }).ToList();

            List<UserTextEntry> gridData;
            if (_aggregateCheckBox.Checked == true)
            {
                var aggregatedData = processedObjects
                    .GroupBy(p => new { p.Material, p.IfcClass, p.QuantityType })
                    .Select(g =>
                    {
                        var totalQuantity = g.Sum(p => p.Quantity);
                        var totalLca = 0.0;
                        _materialLcaData.TryGetValue(g.Key.Material, out MaterialData materialData);
                        if (materialData != null && g.Key.QuantityType != "N/A") totalLca = totalQuantity * materialData.Lca;

                        return new UserTextEntry
                        {
                            Value = g.Key.Material,
                            IfcClass = g.Key.IfcClass,
                            ReferenceQuantity = materialData?.ReferenceQuantity ?? 0,
                            ReferenceUnit = materialData?.ReferenceUnit ?? "",
                            ReferenceLca = materialData?.Lca ?? 0,
                            Count = g.Count(),
                            Quantity = totalQuantity,
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
            _totalLcaLabel.Text = $"Total LCA: {gridData.Sum(entry => entry.Lca):0.00} kgCO2 eq";

            SyncGridSelectionWithRhino();
        }

        private void ReloadIfcClassList()
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;
            var newIfcClasses = CsvReader.ReadIfcClassesDynamic(doc);
            if (!newIfcClasses.SequenceEqual(_ifcClasses))
            {
                _ifcClasses = newIfcClasses;
                PopulateIfcDropdown();
            }
        }

        private void PopulateIfcDropdown()
        {
            var selectedText = _ifcDropdown.SelectedValue is ListItem selectedItem ? selectedItem.Text : null;
            _ifcDropdown.Items.Clear();
            _ifcDropdown.Items.Add(new ListItem { Text = "" });
            foreach (var cls in _ifcClasses) _ifcDropdown.Items.Add(new ListItem { Text = cls });
            if (selectedText != null)
            {
                var itemToRestore = _ifcDropdown.Items.FirstOrDefault(i => i.Text == selectedText);
                if (itemToRestore != null) _ifcDropdown.SelectedValue = itemToRestore;
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
            var keyCount = docStrings.Count;
            if (keyCount == 0)
            {
                _docUserTextGridView.DataStore = null;
            }
            else
            {
                var data = new List<DocumentUserTextEntry>(keyCount);
                for (int i = 0; i < keyCount; i++)
                {
                    var key = docStrings.GetKey(i);
                    if (!string.IsNullOrEmpty(key)) data.Add(new DocumentUserTextEntry { Key = key, Value = docStrings.GetValue(key) });
                }
                _docUserTextGridView.DataStore = data.OrderBy(d => d.Key).ToList();
            }
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
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;
            var value = selectedItem?.Text;
            if (string.IsNullOrWhiteSpace(value))
            {
                RhinoApp.WriteLine($"Please select a valid {friendlyName} from the dropdown first.");
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
                case Brep brep when brep.IsSolid:
                    quantity = brep.GetVolume();
                    break;
                case Extrusion extrusion:
                    var brepFromExtrusion = extrusion.ToBrep();
                    if (brepFromExtrusion != null && brepFromExtrusion.IsSolid)
                        quantity = brepFromExtrusion.GetVolume();
                    break;
                case Mesh mesh when mesh.IsClosed:
                    quantity = mesh.Volume();
                    break;
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
        #endregion
    }
}