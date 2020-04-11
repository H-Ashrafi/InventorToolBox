using Autodesk.iLogic.Interfaces;
using Inventor;
namespace InventorToolBox
{
    public class DocumentManager :IManager
    {
        public Document GetDocument(Application InventorApplicaiton, string fullFileName, bool openVisible = true)
        {
            if (System.IO.File.Exists(fullFileName))
                return InventorApplicaiton.Documents.Open(fullFileName, openVisible);
            return null;
        }
    }
}