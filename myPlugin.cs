using Autodesk.AutoCAD.Runtime;

[assembly: ExtensionApplication(typeof(AutoCADExcel.MyPlugin))]

namespace AutoCADExcel
{
    public class MyPlugin : IExtensionApplication
    {
        void IExtensionApplication.Initialize()
        {
            MyCommands.MakeRibbon();
        }

        void IExtensionApplication.Terminate()
        {
        }

    }

}
