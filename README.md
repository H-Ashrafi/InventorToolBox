# Inventor Toolbox OverView 🧰
This is a library of usefull extension and helper methodes for Autodesk Inventor API. 

|These Inventor Interfaces are covered|
|----------------------------------|
|Application|
|AssemblyDocument|
|ComponentOccurances|
|Documen|
|PartDocument|

# Samlpe
```csharp
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
            App.ConnectToInventor();

            //Get partNo of active document
            var partNo = App.ActiveDocument.GetProperty(kDocumnetProperty.PartNumber);

            //cast into string...
            if (partNo.Value is string value)

                //return value if successful
                return value;

            //return not found if un-successful
            return "not found";
        }
    }
```
## Language
C# 
