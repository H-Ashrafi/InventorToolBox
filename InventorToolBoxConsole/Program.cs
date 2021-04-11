using static InventorToolBox.App;
using InventorToolBox;
using Inventor;
using System;

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
            Console.WriteLine(GetQty().ToString());
            Console.ReadLine();
        }

        private static int GetQty()
        {
            var bomManager = new AssemblyDocumentExtensions();
            var docManager = new DocumentManager();
            var assembly = docManager.GetDocument(InventorApp, @"C:\CAD\Designs\0143-GMW-2019-10 - Yarrawonga Spillway\0143-GMW-2019-1000 - NORTH WEIR\MODULE 1\STAIRWAY_1\STAIRWAY.iam");
            var target =   docManager.GetDocument(InventorApp, @"C:\CAD\Designs\0143-GMW-2019-10 - Yarrawonga Spillway\0143-GMW-2019-1000 - NORTH WEIR\MODULE 1\STAIRWAY_1\Mebmers of STRINGER\STRINGER_MEMBER_001.ipt");
            var target2 = docManager.GetDocument(InventorApp, @"C:\CAD\Designs\0143-GMW-2019-10 - Yarrawonga Spillway\0143-GMW-2019-1000 - NORTH WEIR\MODULE 1\STAIRWAY_1\skeleton.ipt");
            var target3 = docManager.GetDocument(InventorApp, @"C:\CAD\Designs\0143-GMW-2019-10 - Yarrawonga Spillway\0143-GMW-2019-1000 - NORTH WEIR\MODULE 1\STAIRWAY_1\GUARDRAIL.iam");
         
            return bomManager.GetPartsOnlyQuantity(target2, (AssemblyDocument)assembly,true);
        }

        private static void TestPropertyManager()
        {
            var propetyManager = (iPropertiesManager)GetManager(kManagerTypes.iProperties);
            propetyManager.SetCustomProperty(ActiveDocument, "MycustomPropert", 10);
            propetyManager.SetCustomProperty(ActiveDocument, "MycustomPropert2", true);
        }
    }
}


