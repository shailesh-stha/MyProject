// MyProjectPanelCommand.cs
using Rhino;
using Rhino.Commands;
using Rhino.UI;
using System;
using System.Runtime.InteropServices;

namespace MyProject
{
    [Guid("D86F9756-3A83-42D6-BC64-16CF585D7576")]
    public class MyProjectPanelCommand : Command
    {
        public MyProjectPanelCommand() { }

        public override string EnglishName => "MyProjectPanelCommand";

        protected override Result RunCommand(RhinoDoc doc, RunMode mode)
        {
            const string panelId = "01D22D46-3C89-43A5-A92B-A1D6519A4244";
            Panels.OpenPanel(new Guid(panelId));
            return Result.Success;
        }
    }
}