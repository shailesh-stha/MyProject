// MyProjectPlugin.cs
using Rhino.PlugIns;
using Rhino.UI;

namespace MyProject
{
    public class MyProjectPlugin : PlugIn
    {
        public MyProjectPlugin()
        {
            Instance = this;
        }

        public static MyProjectPlugin Instance { get; private set; }

        protected override LoadReturnCode OnLoad(ref string errorMessage)
        {
            Panels.RegisterPanel(this, typeof(MyProjectPanel), "My Project Panel", null);
            return LoadReturnCode.Success;
        }
    }
}
