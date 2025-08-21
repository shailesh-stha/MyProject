// MyProjectPlugin.cs
using Rhino;
using Rhino.UI;
using Rhino.PlugIns;
using System;
using System.Drawing;
using System.IO; // Required for MemoryStream

namespace MyProject
{
    ///<summary>
    /// <para>Every RhinoCommon .rhp assembly must have one and only one PlugIn-derived
    /// class. DO NOT create a second class with the PlugIn attribute.</para>
    /// <para>Instead, load another .dll from your PlugIn.OnLoad method.</para>
    ///</summary>
    public class MyProjectPlugin : Rhino.PlugIns.PlugIn
    {
        public MyProjectPlugin()
        {
            Instance = this;
        }

        public static MyProjectPlugin Instance { get; private set; }

        protected override LoadReturnCode OnLoad(ref string errorMessage)
        {
            // Convert the byte array from resources into an Icon object
            Icon panelIcon = BytesToIcon(Properties.Resources.icon_strLCA32);

            // Pass the newly created Icon object to the RegisterPanel method
            Panels.RegisterPanel(this, typeof(MyProjectPanel), "◀◀ StrLcaPro ▶▶", panelIcon);
            return LoadReturnCode.Success;
        }

        protected override void OnShutdown()
        {
            // Panel cleanup, if necessary
        }

        /// <summary>
        /// Helper method to convert a byte array into a System.Drawing.Icon
        /// </summary>
        /// <param name="bytes">The byte array containing the icon data.</param>
        /// <returns>A System.Drawing.Icon object.</returns>
        private Icon BytesToIcon(byte[] bytes)
        {
            using (var ms = new MemoryStream(bytes))
            {
                return new Icon(ms);
            }
        }
    }
}
