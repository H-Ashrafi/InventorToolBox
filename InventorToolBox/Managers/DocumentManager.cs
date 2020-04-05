using Autodesk.iLogic.Interfaces;
using Inventor;
namespace InventorToolBox
{
    public static class DocumentManager 
    {
        public static Document GetDocument(Application InventorApplicaiton, string fullFileName, bool openVisible = true)
        {
            if (System.IO.File.Exists(fullFileName))
                return InventorApplicaiton.Documents.Open(fullFileName, openVisible);
            return null;
        }
    }
}