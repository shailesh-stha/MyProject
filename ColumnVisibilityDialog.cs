// ColumnVisibilityDialog.cs
using Eto.Forms;
using Eto.Drawing;
using System.Collections.Generic;

namespace MyProject
{
    public class ColumnVisibilityDialog : Dialog
    {
        private readonly GridView _gridView;

        public ColumnVisibilityDialog(GridView gridView)
        {
            Title = "Show/Hide Columns";
            MinimumSize = new Size(200, 200);
            _gridView = gridView;

            var layout = new DynamicLayout();
            layout.BeginVertical(new Padding(10), new Size(5, 5));

            foreach (var column in _gridView.Columns)
            {
                var checkbox = new CheckBox { Text = column.HeaderText, Checked = column.Visible };

                checkbox.CheckedChanged += (s, e) =>
                {
                    column.Visible = checkbox.Checked ?? false;
                };

                layout.Add(checkbox);
            }
            layout.EndVertical();

            Content = layout;
        }
    }
}