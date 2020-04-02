using System;
using Inventor;
namespace InventorToolBox
{
    public class App
    {
        private readonly static object _lock = new object();
        private static Application _Inventor;
        private void SetApplication()
        {
            try
            {
                try
                {
                    _Inventor = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application")
                        as Inventor.Application;
                }
                catch
                {
                    Type inventorAppType = System.Type.GetTypeFromProgID("Inventor.Application");
                    var app = System.Activator.CreateInstance(inventorAppType) as Inventor.Application;
                    //Must be set visible explicitly
                    app.Visible = true;
                    _Inventor = app;
                }
            }
            catch (Exception)
            {
                throw new MemberAccessException("Could not get an instance of Inventor");
            }

        }
        private static App _Instance;
        public static App ConnectToInventor()
        {
            lock (_lock)
            {
                if (_Instance == null)
                {
                    _Instance = new App();
                    _Instance.SetApplication();
                }
                return _Instance;
            }
        }
        public static Application InventorApp=>_Inventor;
        public static IManager GetManager (kManagerTypes mangerType)
        {
            switch (mangerType)
            {
                case kManagerTypes.Document:return new DocumentManager();
                case kManagerTypes.Part: return new PartManager();
                case kManagerTypes.Assembly: return new AssemblyManager();
                case kManagerTypes.Drawing:return new DrawingManager();
                case kManagerTypes.Presentation:return new PresentationManger();
                case kManagerTypes.iProperties:return new iPropertiesManager();
                case kManagerTypes.BOM:return new BomManager();
                default: return null;
            }
        }
        public static Document ActiveDocument => _Inventor.ActiveDocument;
    }
}
