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
using InventorToolBox;
//set up a console app
//get an instance of Inventor
            App.ConnectToInventor();

            //Get partNo of active document
            var partNo = App.ActiveDocument.GetProperty(kDocumnetProperty.PartNumber);
```
[.CS file here](https://github.com/H-Ashrafi/InventorToolBox/blob/master/InventorToolBoxConsole/Program.cs)
## Language
C# 
