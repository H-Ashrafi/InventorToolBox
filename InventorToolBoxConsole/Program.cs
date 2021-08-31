using Inventor;
using InventorToolBox;
using System;

namespace InventorToolBoxConsole
{
    /// <summary>
    /// a console app as a sample applicaiton for InventorToolbox
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            //message to user
            Console.WriteLine("Part Number of active document:");

            //run the funcions
            Console.WriteLine(GetPartNumberOfActiveDocument());
            Console.ReadLine();
        }

        /// <summary>
        /// get the part number of active document
        /// </summary>
        /// <returns>returns part number of active doc as string and empty string if not of type string</returns>
        private static string GetPartNumberOfActiveDocument()
        {
            //get an instance of Inventor
            var app = App.ConnectToInventor();
            var part = app.NewPart();
            var sketch = part.AddSketch(KMainPlane.xy);

            //Get partNo of active document
            var partNo = part.AsDocument().GetProperty(kDocumnetProperty.PartNumber);

            //cast into string...
            if (partNo.Value is string value)

                //return value if successful
                return value;

            //return not found if un-successful
            return "not found";
        }
    }
}


