// MyProjectPanelCommand.cs
using Rhino;
using Rhino.Commands;
using Rhino.UI;

namespace MyProject
{
    public class MyProjectPanelCommand : Command
    {
        public MyProjectPanelCommand()
        {
            Instance = this;
        }

        public static MyProjectPanelCommand Instance { get; private set; }

        public override string EnglishName => "MyProjectPanel";

        protected override Result RunCommand(RhinoDoc doc, RunMode mode)
        {
            var panelGuid = typeof(MyProjectPanel).GUID;

            if (!Panels.IsPanelVisible(panelGuid))
                Panels.OpenPanel(panelGuid);
            else
                Panels.ClosePanel(panelGuid);

            return Result.Success;
        }
    }
}
