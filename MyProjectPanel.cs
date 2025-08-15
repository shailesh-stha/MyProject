// MyProjectPanel.cs
using Rhino;
using Rhino.Geometry;
using Rhino.Input.Custom;
using Rhino.DocObjects;
using Eto.Forms;
using Eto.Drawing;
using System;
using System.Collections.Generic;
using System.Diagnostics; // Needed for Process.Start
using System.Linq;
using System.Runtime.InteropServices;
using Rhino.UI.Controls;

namespace MyProject
{
    [Guid("01D22D46-3C89-43A5-A92B-A1D6519A4244")]
    public class MyProjectPanel : Panel
    {
        public static string SelectedMaterial { get; private set; }

        // Alphabetical materials list
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

        // UI elements updated dynamically
        private Label _volumeDescriptionLabel;
        private Label _volumeValueLabel;

        public MyProjectPanel(uint documentSerialNumber)
        {
            // Create a bold font for section titles
            var defaultFont = Eto.Drawing.SystemFonts.Default();
            var boldFont = new Eto.Drawing.Font(defaultFont.Family, defaultFont.Size, Eto.Drawing.FontStyle.Bold);

            // ===== Top section: buttons =====
            var btnPickPoint = new Button { Text = "Pick Point", Command = new PickPointCommand() };
            var btnDrawRect = new Button { Text = "Draw Rectangle", Command = new DrawRectangleCommand() };

            var btnStructure = new Button { Text = "str-ucture" };
            btnStructure.Click += (s, e) =>
            {
                var url = "http://www.str-ucture.com";
                try
                {
                    Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    RhinoApp.WriteLine($"Error opening website: {ex.Message}");
                }
            };

            // UPDATED: Use a horizontal StackLayout for the buttons to control spacing
            var buttonLayout = new StackLayout
            {
                Orientation = Orientation.Horizontal,
                Spacing = 5, // This adds a 5-pixel gap between each button
                Items = { btnPickPoint, btnDrawRect, btnStructure }
            };

            var upperLayout = new DynamicLayout { Spacing = new Size(6, 6), Padding = new Padding(10) };
            upperLayout.AddRow(new Label { Text = "The Buttons", Font = boldFont });
            // UPDATED: Add the new button layout to the panel
            upperLayout.AddRow(buttonLayout);
            upperLayout.Add(null); // spacer

            // ===== Middle (second) section: dropdown =====
            var dropdown = new ComboBox();
            dropdown.Items.Add(new ListItem { Text = "" }); // allow clearing selection
            foreach (var m in _materials) dropdown.Items.Add(new ListItem { Text = m });

            dropdown.SelectedIndexChanged += (s, e) =>
            {
                var li = dropdown.SelectedValue as ListItem;
                var val = li?.Text;
                if (string.IsNullOrWhiteSpace(val))
                {
                    SelectedMaterial = null;
                    RhinoApp.WriteLine("Material selection cleared.");
                    return;
                }

                SelectedMaterial = val;
                RhinoApp.WriteLine($"Material selected: {SelectedMaterial}");
            };

            var middleLayout = new DynamicLayout { Spacing = new Size(6, 6), Padding = new Padding(10) };
            middleLayout.AddRow(new Divider());
            middleLayout.AddRow(new Label { Text = "The Dropdown", Font = boldFont });
            middleLayout.AddRow(new Label { Text = "Select Material:" }, dropdown);
            middleLayout.Add(null); // spacer

            // ===== Bottom (third) section: selection info (volume) =====
            _volumeValueLabel = new Label
            {
                Text = "Select one or more closed solids to display volume.",
                TextColor = Color.FromGrayscale(0.45f)
            };

            _volumeDescriptionLabel = new Label { Text = "Selected Object Volume:" };

            var bottomLayout = new DynamicLayout { Spacing = new Size(6, 6), Padding = new Padding(10) };
            bottomLayout.AddRow(new Divider());
            bottomLayout.AddRow(new Label { Text = "The Calculation", Font = boldFont });
            bottomLayout.AddRow(_volumeDescriptionLabel);
            bottomLayout.AddRow(_volumeValueLabel);
            bottomLayout.Add(null); // spacer

            // Stack middle + bottom into one container for the Splitter's Panel2
            var lowerContainer = new StackLayout
            {
                Orientation = Orientation.Vertical,
                Spacing = 0,
                Items = { middleLayout, bottomLayout }
            };

            // Splitter: top=buttons, bottom=middle+bottom
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
            RhinoDoc.SelectObjects += OnSelectObjects;
            RhinoDoc.DeselectObjects += OnDeselectObjects;
            RhinoDoc.DeselectAllObjects += OnDeselectAllObjects;
            RhinoDoc.AddRhinoObject += OnAddOrDeleteObject;
            RhinoDoc.DeleteRhinoObject += OnAddOrDeleteObject;
            RhinoDoc.ReplaceRhinoObject += OnReplaceObject;

            // Initial update
            UpdateSelectedObjectVolume();
        }

        protected override void Dispose(bool disposing)
        {
            // Unsubscribe to avoid leaks
            RhinoDoc.SelectObjects -= OnSelectObjects;
            RhinoDoc.DeselectObjects -= OnDeselectObjects;
            RhinoDoc.DeselectAllObjects -= OnDeselectAllObjects;
            RhinoDoc.AddRhinoObject -= OnAddOrDeleteObject;
            RhinoDoc.DeleteRhinoObject -= OnAddOrDeleteObject;
            RhinoDoc.ReplaceRhinoObject -= OnReplaceObject;

            base.Dispose(disposing);
        }

        // ---- Event handlers ----
        private void OnSelectObjects(object sender, RhinoObjectSelectionEventArgs e) => UpdateSelectedObjectVolume();
        private void OnDeselectObjects(object sender, RhinoObjectSelectionEventArgs e) => UpdateSelectedObjectVolume();
        private void OnDeselectAllObjects(object sender, RhinoDeselectAllObjectsEventArgs e) => UpdateSelectedObjectVolume();
        private void OnAddOrDeleteObject(object sender, RhinoObjectEventArgs e) => UpdateSelectedObjectVolume();
        private void OnReplaceObject(object sender, RhinoReplaceObjectEventArgs e) => UpdateSelectedObjectVolume();

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

            // Case 0: Nothing selected
            if (selectionCount == 0)
            {
                _volumeDescriptionLabel.Text = "Selected Object Volume:";
                SetVolumeLabel("Nothing selected.", isInfo: true);
                return;
            }

            // Case 1: Exactly one object is selected
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

            // Case 2: Multiple objects are selected
            if (selectionCount > 1)
            {
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
        }

        private void SetVolumeLabel(string message, bool isInfo)
        {
            _volumeValueLabel.Text = message;
            _volumeValueLabel.TextColor = isInfo ? Color.FromGrayscale(0.45f) : Colors.Black;
        }

        // --- Volume helpers (unchanged) ---
        private bool TryComputeVolume(GeometryBase geo, out double volume)
        {
            volume = 0.0;
            if (geo == null) return false;

            try
            {
                switch (geo)
                {
                    case Brep brep:
                        if (!brep.IsSolid) return false;
                        using (var vmp = VolumeMassProperties.Compute(brep, true, true, true, true))
                        {
                            if (vmp == null) return false;
                            volume = vmp.Volume;
                            return true;
                        }

                    case Extrusion extrusion:
                        {
                            var b = extrusion.ToBrep();
                            if (b == null || !b.IsSolid) return false;
                            using (var vmp = VolumeMassProperties.Compute(b, true, true, true, true))
                            {
                                if (vmp == null) return false;
                                volume = vmp.Volume;
                                return true;
                            }
                        }

                    case Mesh mesh:
                        if (!mesh.IsClosed) return false;
                        using (var vmp = VolumeMassProperties.Compute(mesh))
                        {
                            if (vmp == null) return false;
                            volume = vmp.Volume;
                            return true;
                        }

                    default:
                        var asBrep = Brep.TryConvertBrep(geo);
                        if (asBrep != null && asBrep.IsSolid)
                        {
                            using (var vmp = VolumeMassProperties.Compute(asBrep, true, true, true, true))
                            {
                                if (vmp == null) return false;
                                volume = vmp.Volume;
                                return true;
                            }
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Volume compute error: {ex.Message}");
                return false;
            }

            return false;
        }
    }

    // ===== Commands (unchanged) =====

    internal class PickPointCommand : Command
    {
        public PickPointCommand() => Executed += OnExecuted;

        private void OnExecuted(object sender, EventArgs e)
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;

            var gp = new GetPoint();
            gp.SetCommandPrompt("Pick a point");
            var res = gp.Get();
            if (res != Rhino.Input.GetResult.Point || gp.CommandResult() != Rhino.Commands.Result.Success)
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
            var res1 = gp1.Get();
            if (res1 != Rhino.Input.GetResult.Point || gp1.CommandResult() != Rhino.Commands.Result.Success)
                return;

            var p1 = gp1.Point();

            var gp2 = new GetPoint();
            gp2.SetCommandPrompt("Opposite corner of rectangle");
            gp2.ConstrainToConstructionPlane(true);
            gp2.DrawLineFromPoint(p1, true);

            gp2.DynamicDraw += (s, args) =>
            {
                var current = args.CurrentPoint;
                var previewRect = new Rectangle3d(plane, p1, current);
                var pl = previewRect.ToPolyline();
                if (pl != null)
                    args.Display.DrawPolyline(pl, System.Drawing.Color.White, 2);
            };

            var res2 = gp2.Get();
            if (res2 != Rhino.Input.GetResult.Point || gp2.CommandResult() != Rhino.Commands.Result.Success)
                return;

            var p2 = gp2.Point();
            var rect = new Rectangle3d(plane, p1, p2);
            var curve = rect.ToNurbsCurve();

            if (curve != null)
            {
                doc.Objects.AddCurve(curve);
                doc.Views.Redraw();
                RhinoApp.WriteLine("Rectangle created.");
            }
        }
    }
}