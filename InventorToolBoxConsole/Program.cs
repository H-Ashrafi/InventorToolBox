using static InventorToolBox.App;
using InventorToolBox;
namespace InventorToolBoxConsole
{
    class Program
    {
        //public static AutoInvent app = AutoInvent.Initiate;
        static void Main(string[] args)
        {
            TestAPI();
        }

        private static void TestAPI()
        {
            ConnectToInventor();
            var propetyManager=(iPropertiesManager)GetManager(kManagerTypes.iProperties);
            propetyManager.SetCustomProperty(ActiveDocument, "MycustomPropert", 10);
            propetyManager.SetCustomProperty(ActiveDocument, "MycustomPropert2", true);
        }
    }
}


