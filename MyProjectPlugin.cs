// MyProjectPlugin.cs
using Rhino.PlugIns;
using Rhino.UI;
using System;

namespace MyProject
{
    ///<summary>
    /// <para>Every RhinoCommon .rhp assembly must have one and only one PlugIn-derived
    /// class. DO NOT create a second class with the PlugIn attribute.</para>
    /// <para>Instead, load another .dll from your PlugIn.OnLoad method.</para>
    ///</summary>
    public class MyProjectPlugin : PlugIn
    {
        public MyProjectPlugin()
        {
            Instance = this;
        }

        ///<summary>Gets the only instance of the MyProjectPlugin plug-in.</summary>
        public static MyProjectPlugin Instance { get; private set; }

        protected override LoadReturnCode OnLoad(ref string errorMessage)
        {
            Panels.RegisterPanel(this, typeof(MyProjectPanel), "MyProjectPanel", null);
            return LoadReturnCode.Success;
        }

        protected override void OnShutdown()
        {
            // Panel cleanup, if necessary
        }
    }
}