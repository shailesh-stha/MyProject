// SettingsDialog.cs
using Eto.Forms;
using Eto.Drawing;

namespace MyProject
{
    /// <summary>
    /// A simple dialog window for plugin settings.
    /// This serves as a placeholder for future configuration options.
    /// </summary>
    public class SettingsDialog : Dialog
    {
        public SettingsDialog()
        {
            // --- Dialog Properties ---
            Title = "Plugin Settings";
            MinimumSize = new Size(350, 200);
            Padding = new Padding(10);

            // --- Controls ---
            var testLabel = new Label
            {
                Text = "this is test settings",
                VerticalAlignment = VerticalAlignment.Center,
                // --- MODIFICATION: Removed the invalid HorizontalAlignment property ---
            };

            // --- Layout ---
            // A simple layout to center the label.
            var layout = new DynamicLayout();
            layout.Add(null); // Spacer
            layout.AddRow(null, testLabel, null); // Center the label horizontally
            layout.Add(null); // Spacer

            Content = layout;

            // --- Buttons ---
            // Add a default "OK" button to close the dialog.
            DefaultButton = new Button { Text = "OK" };
            DefaultButton.Click += (s, e) => Close();

            // Optional: Add an AbortButton to handle the Escape key.
            AbortButton = DefaultButton;
        }
    }
}