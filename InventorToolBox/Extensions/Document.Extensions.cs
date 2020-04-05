using Inventor;
namespace InventorToolBox
{
    public static class DocumentExtensions
    {
        /// <summary>
        /// creates a new part document
        /// </summary>
        /// <returns></returns>
        public static Document AddNewPart(this Application App, string TemplateFileName="",bool CreateVisible=true)
        {
           return App.Documents.Add(DocumentTypeEnum.kPartDocumentObject, TemplateFileName, CreateVisible);
        }

        /// <summary>
        /// opens an inventor document
        /// </summary>
        /// <param name="FullFileName">full file name of the inventor file</param>
        /// <param name="OpenVisible">set to false to do it silently</param>
        /// <returns>opens the file in the active instance of inventor and returns the active document if an error occured</returns>
        public static Document OpenInventorDoc(this Application App, string FullFileName, bool OpenVisible = true)
        {
            try
            {
                return App.Documents.Open(FullFileName, OpenVisible);
            }
            catch
            {
                return App.ActiveDocument;
            }
        }

        /// <summary>
        /// get the active drawing document
        /// </summary>
        /// <param name="App">Inventor.Application object</param>
        /// <returns></returns>
        public static DrawingDocument GetDrawingDoc(this Application App)
        {
            if (App.ActiveDocumentType == DocumentTypeEnum.kDrawingDocumentObject)
            {
                return App.ActiveDocument as DrawingDocument;
            }
            return null;
        }

        /// <summary>
        /// get the active assembly document
        /// </summary>
        /// <param name="App">Inventor.Application object</param>
        /// <returns></returns>
        public static AssemblyDocument GetAssemblyDoc(this Application App)
        {
            if (App.ActiveDocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                return App.ActiveDocument as AssemblyDocument;
            }
            return null;
        }

        /// <summary>
        /// get inventor active part document
        /// </summary>
        /// <param name="App">inventor.application object</param>
        /// <returns></returns>
        public static PartDocument GetPartDoc(this Application App)
        {
            if (App.ActiveDocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                return App.ActiveDocument as PartDocument;
            }
            return null;
        }
    }
}
