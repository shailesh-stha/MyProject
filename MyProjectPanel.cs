// MyProjectPanel.cs
using Rhino;
using Rhino.Geometry;
using Rhino.Input.Custom;
using Eto.Forms;
using Eto.Drawing;
using System;
using System.Collections.Generic;
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
            "Concrete",
            "Structural Steel",
            "Timber",
            "Aluminum",
            "Masonry",
            "Glass",
            "Cast Iron",
            "Composite Materials",
            "Reinforced Concrete",
            "Bamboo"
        };

        public MyProjectPanel(uint documentSerialNumber)
        {
            var sortedMaterials = _materials.OrderBy(m => m).ToList();

            // ===== Top section: buttons =====
            var btnPickPoint = new Button { Text = "Pick Point", Command = new PickPointCommand() };
            var btnDrawRect = new Button { Text = "Draw Rectangle", Command = new DrawRectangleCommand() };

            var upperLayout = new DynamicLayout { Spacing = new Size(6, 6), Padding = new Padding(10) };
            upperLayout.AddSeparateRow(btnPickPoint, btnDrawRect, null);
            upperLayout.Add(null); // spacer

            // ===== Bottom section =====
            var dropdown = new ComboBox();
            dropdown.Items.Add(new ListItem { Text = "" });
            foreach (var m in sortedMaterials)
                dropdown.Items.Add(new ListItem { Text = m });

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

            var lowerLayout = new DynamicLayout { Spacing = new Size(6, 6), Padding = new Padding(10) };

            // Custom separator (1px high panel with light gray color)
            var separator = new Panel
            {
                BackgroundColor = Color.FromGrayscale(0.85f),
                Height = 1
            };
            lowerLayout.AddRow(separator);

            lowerLayout.AddRow(new Label { Text = "Select Material:" }, dropdown);
            lowerLayout.Add(null); // spacer

            var splitter = new Splitter
            {
                Orientation = Orientation.Vertical,
                Panel1 = upperLayout,
                Panel2 = lowerLayout,
                FixedPanel = SplitterFixedPanel.Panel1,
                Position = 120
            };

            Content = splitter;
            Size = new Size(440, 600);
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
