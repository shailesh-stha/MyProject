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

namespace MyProject
{
    [Guid("01D22D46-3C89-43A5-A92B-A1D6519A4244")]
    public class MyProjectPanel : Panel
    {
        public static string SelectedMaterial { get; private set; }

        private readonly List<string> _materials = new List<string>
        {
            "Aluminum",
            "Bamboo",
            "Cast Iron",
            "Composite Materials",
            "Concrete",
            "Glass",
            "Masonry",
            "Reinforced Concrete",
            "Structural Steel",
            "Timber"
        };

        private Label _volumeDescriptionLabel;
        private Label _volumeValueLabel;

        private class UserTextEntry
        {
            public int SerialNumber { get; set; }
            public string Key { get; set; }
            public string Value { get; set; }
            public int Count { get; set; }
        }

        private readonly GridView<UserTextEntry> _userTextGridView;

        public MyProjectPanel(uint documentSerialNumber)
        {
            var defaultFont = Eto.Drawing.SystemFonts.Default();
            var boldFont = new Eto.Drawing.Font(defaultFont.Family, defaultFont.Size, Eto.Drawing.FontStyle.Bold);

            // ===== Top section: Buttons for testing =====
            var btnPickPoint = new Button { Text = "Pick Point", Command = new PickPointCommand() };
            var btnDrawRect = new Button { Text = "Draw Rectangle", Command = new DrawRectangleCommand() };
            var btnStructure = new Button { Text = "str-ucture" };
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
                Items = { btnPickPoint, btnDrawRect, btnStructure }
            };

            var upperLayout = new DynamicLayout { Spacing = new Size(6, 6), Padding = new Padding(10) };
            upperLayout.AddRow(new Label { Text = "Buttons for testing", Font = boldFont });
            upperLayout.AddRow(buttonLayout);
            upperLayout.Add(null);

            // ===== Middle section: Material Defination =====
            var dropdown = new ComboBox();
            dropdown.Items.Add(new ListItem { Text = "" });
            foreach (var m in _materials) dropdown.Items.Add(new ListItem { Text = m });

            dropdown.SelectedIndexChanged += (s, e) =>
            {
                var li = dropdown.SelectedValue as ListItem;
                var val = li?.Text;
                SelectedMaterial = string.IsNullOrWhiteSpace(val) ? null : val;
                RhinoApp.WriteLine(SelectedMaterial == null ? "Material selection cleared." : $"Material selected: {SelectedMaterial}");
            };

            var assignButton = new Button { Text = "Assign" };
            assignButton.Click += OnAssignButtonClick;

            var middleLayout = new DynamicLayout { Spacing = new Size(6, 6), Padding = new Padding(10) };
            middleLayout.AddRow(new Divider());
            middleLayout.AddRow(new Label { Text = "Material Defination", Font = boldFont });
            middleLayout.AddRow(new Label { Text = "Select Material:" }, dropdown, assignButton);
            middleLayout.Add(null);

            // ===== User Text Section =====
            _userTextGridView = new GridView<UserTextEntry>
            {
                ShowHeader = true,
                AllowMultipleSelection = false,
                Columns =
                {
                    new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.SerialNumber)), HeaderText = "SN", Editable = false, Width = 40 },
                    new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.Key)), HeaderText = "Key", Editable = false, Width = 120},
                    new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.Value)), HeaderText = "Value", Editable = false, AutoSize = true},
                    new GridColumn { DataCell = new TextBoxCell(nameof(UserTextEntry.Count)), HeaderText = "Count", Editable = false, Width = 60 }
                }
            };

            var userTextLayout = new DynamicLayout { Spacing = new Size(6, 6), Padding = new Padding(10) };
            userTextLayout.AddRow(new Divider());
            userTextLayout.AddRow(new Label { Text = "Attribute User Text", Font = boldFont });
            userTextLayout.AddRow(_userTextGridView);
            userTextLayout.Add(null);

            // ===== Bottom section: Values and Calculation =====
            _volumeValueLabel = new Label
            {
                Text = "Select one or more closed solids to display volume.",
                TextColor = Color.FromGrayscale(0.45f)
            };

            _volumeDescriptionLabel = new Label { Text = "Selected Object Volume:" };

            var bottomLayout = new DynamicLayout { Spacing = new Size(6, 6), Padding = new Padding(10) };
            bottomLayout.AddRow(new Divider());
            bottomLayout.AddRow(new Label { Text = "Values and Calculation", Font = boldFont });
            bottomLayout.AddRow(_volumeDescriptionLabel);
            bottomLayout.AddRow(_volumeValueLabel);
            bottomLayout.Add(null);

            var lowerContainer = new StackLayout
            {
                Orientation = Orientation.Vertical,
                Spacing = 0,
                Items = { middleLayout, userTextLayout, bottomLayout }
            };

            var splitter = new Splitter
            {
                Orientation = Orientation.Vertical,
                Panel1 = upperLayout,
                Panel2 = lowerContainer,
                FixedPanel = SplitterFixedPanel.Panel1,
                Position = 140
            };

            Content = splitter;
            Size = new Size(440, 600);

            // ---- Rhino document event hooks ----
            RhinoDoc.SelectObjects += OnSelectionChanged;
            RhinoDoc.DeselectObjects += OnSelectionChanged;
            RhinoDoc.DeselectAllObjects += OnDeselectAllObjects;
            RhinoDoc.AddRhinoObject += OnDatabaseChanged;
            RhinoDoc.DeleteRhinoObject += OnDatabaseChanged;
            RhinoDoc.ReplaceRhinoObject += OnReplaceObject;
            RhinoDoc.UserStringChanged += OnUserStringChanged;

            UpdatePanelData();
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
                if (doc.Objects.ModifyAttributes(rhinoObject.Id, newAttributes, false))
                {
                    updatedCount++;
                }
            }

            doc.Views.Redraw();
            RhinoApp.WriteLine($"Successfully assigned '{SelectedMaterial}' to {updatedCount} object(s).");
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
            base.Dispose(disposing);
        }

        private void OnSelectionChanged(object sender, RhinoObjectSelectionEventArgs e) => UpdatePanelData();
        private void OnDeselectAllObjects(object sender, RhinoDeselectAllObjectsEventArgs e) => UpdatePanelData();
        private void OnDatabaseChanged(object sender, RhinoObjectEventArgs e) => UpdatePanelData();
        private void OnReplaceObject(object sender, RhinoReplaceObjectEventArgs e) => UpdatePanelData();
        private void OnUserStringChanged(object sender, RhinoDoc.UserStringChangedArgs e) => UpdatePanelData();

        private void UpdatePanelData()
        {
            Eto.Forms.Application.Instance.AsyncInvoke(() =>
            {
                UpdateSelectedObjectVolume();
                UpdateUserTextGrid();
            });
        }

        private void UpdateUserTextGrid()
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;

            var selectedObjects = doc.Objects.GetSelectedObjects(false, false).ToList();
            if (selectedObjects.Count == 0)
            {
                _userTextGridView.DataStore = new List<UserTextEntry>();
                return;
            }

            var pairCounts = new Dictionary<(string Key, string Value), int>();
            int blankObjectCount = 0;

            foreach (var rhinoObject in selectedObjects)
            {
                var userStrings = rhinoObject.Attributes.GetUserStrings();
                if (userStrings == null || userStrings.Count == 0)
                {
                    blankObjectCount++;
                    continue;
                }

                foreach (var key in userStrings.AllKeys)
                {
                    var value = userStrings.Get(key);
                    var pair = (Key: key, Value: value);
                    if (!pairCounts.ContainsKey(pair)) pairCounts[pair] = 0;
                    pairCounts[pair]++;
                }
            }

            if (blankObjectCount > 0)
            {
                pairCounts[("not-assigned", "not-assigned")] = blankObjectCount;
            }

            var gridData = pairCounts.Select(kvp => new UserTextEntry
            {
                Key = kvp.Key.Key,
                Value = kvp.Key.Value,
                Count = kvp.Value
            })
                .OrderBy(entry => entry.Key)
                .ThenBy(entry => entry.Value)
                .ToList();

            for (int i = 0; i < gridData.Count; i++)
            {
                gridData[i].SerialNumber = i + 1;
            }

            _userTextGridView.DataStore = gridData;
        }

        private void UpdateSelectedObjectVolume()
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null)
            {
                SetVolumeLabel("No document.", isInfo: true);
                return;
            }

            var selectedObjects = doc.Objects.GetSelectedObjects(false, false).ToList();
            var selectionCount = selectedObjects.Count;
            var unitAbbreviation = Rhino.UI.Localization.UnitSystemName(doc.ModelUnitSystem, false, true, true);

            if (selectionCount == 0)
            {
                _volumeDescriptionLabel.Text = "Selected Object Volume:";
                SetVolumeLabel("Nothing selected.", isInfo: true);
                return;
            }

            if (selectionCount == 1)
            {
                _volumeDescriptionLabel.Text = "Selected Object Volume:";
                var rhinoObject = selectedObjects[0];
                if (TryComputeVolume(rhinoObject.Geometry, out double volumeValue))
                {
                    _volumeValueLabel.TextColor = Colors.Black;
                    _volumeValueLabel.Text = $"{volumeValue:0.###} {unitAbbreviation}³";
                }
                else
                {
                    SetVolumeLabel("Selected object has no closed volume.", isInfo: true);
                }
                return;
            }

            double totalVolume = 0.0;
            int solidObjectCount = 0;
            foreach (var rhinoObject in selectedObjects)
            {
                if (TryComputeVolume(rhinoObject.Geometry, out double volumeValue))
                {
                    totalVolume += volumeValue;
                    solidObjectCount++;
                }
            }

            if (solidObjectCount > 0)
            {
                _volumeDescriptionLabel.Text = $"Total Volume ({solidObjectCount} Objects):";
                _volumeValueLabel.TextColor = Colors.Black;
                _volumeValueLabel.Text = $"{totalVolume:0.###} {unitAbbreviation}³";
            }
            else
            {
                _volumeDescriptionLabel.Text = "Selected Object Volume:";
                SetVolumeLabel("Selected objects have no closed volume.", isInfo: true);
            }
        }

        private void SetVolumeLabel(string message, bool isInfo)
        {
            _volumeValueLabel.Text = message;
            _volumeValueLabel.TextColor = isInfo ? Color.FromGrayscale(0.45f) : Colors.Black;
        }

        private bool TryComputeVolume(GeometryBase geo, out double volume)
        {
            volume = 0.0;
            if (geo == null) return false;
            try
            {
                switch (geo)
                {
                    case Brep brep when brep.IsSolid:
                        using (var vmp = VolumeMassProperties.Compute(brep, true, true, true, true))
                        {
                            if (vmp != null) { volume = vmp.Volume; return true; }
                        }
                        break;
                    case Extrusion extrusion:
                        var b = extrusion.ToBrep();
                        if (b != null && b.IsSolid)
                        {
                            using (var vmp = VolumeMassProperties.Compute(b, true, true, true, true))
                            {
                                if (vmp != null) { volume = vmp.Volume; return true; }
                            }
                        }
                        break;
                    case Mesh mesh when mesh.IsClosed:
                        using (var vmp = VolumeMassProperties.Compute(mesh))
                        {
                            if (vmp != null) { volume = vmp.Volume; return true; }
                        }
                        break;
                    default:
                        var asBrep = Brep.TryConvertBrep(geo);
                        if (asBrep != null && asBrep.IsSolid)
                        {
                            using (var vmp = VolumeMassProperties.Compute(asBrep, true, true, true, true))
                            {
                                if (vmp != null) { volume = vmp.Volume; return true; }
                            }
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Volume compute error: {ex.Message}");
            }
            return false;
        }
    }

    internal class PickPointCommand : Command
    {
        public PickPointCommand() => Executed += OnExecuted;
        private void OnExecuted(object sender, EventArgs e)
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;
            var gp = new GetPoint();
            gp.SetCommandPrompt("Pick a point");
            if (gp.Get() != Rhino.Input.GetResult.Point || gp.CommandResult() != Rhino.Commands.Result.Success)
                return;
            var pt = gp.Point();
            doc.Objects.AddPoint(pt);
            doc.Views.Redraw();
            RhinoApp.WriteLine($"Point created at {pt}.");
        }
    }

    internal class DrawRectangleCommand : Command
    {
        public DrawRectangleCommand() => Executed += OnExecuted;
        private void OnExecuted(object sender, EventArgs e)
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;
            var view = doc.Views.ActiveView;
            var plane = view?.ActiveViewport?.ConstructionPlane() ?? Plane.WorldXY;

            var gp1 = new GetPoint();
            gp1.SetCommandPrompt("First corner of rectangle");
            gp1.ConstrainToConstructionPlane(true);
            if (gp1.Get() != Rhino.Input.GetResult.Point || gp1.CommandResult() != Rhino.Commands.Result.Success)
                return;
            var p1 = gp1.Point();

            var gp2 = new GetPoint();
            gp2.SetCommandPrompt("Opposite corner of rectangle");
            gp2.ConstrainToConstructionPlane(true);
            gp2.DrawLineFromPoint(p1, true);
            gp2.DynamicDraw += (s, args) =>
            {
                var pl = new Rectangle3d(plane, p1, args.CurrentPoint).ToPolyline();
                if (pl != null) args.Display.DrawPolyline(pl, System.Drawing.Color.White, 2);
            };

            if (gp2.Get() != Rhino.Input.GetResult.Point || gp2.CommandResult() != Rhino.Commands.Result.Success)
                return;
            var p2 = gp2.Point();
            var curve = new Rectangle3d(plane, p1, p2).ToNurbsCurve();

            if (curve != null)
            {
                doc.Objects.AddCurve(curve);
                doc.Views.Redraw();
                RhinoApp.WriteLine("Rectangle created.");
            }
        }
    }
}