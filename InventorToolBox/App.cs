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

        /// <summary>
        /// inventor's <see cref="Application"/> object
        /// </summary>
        public Application InventorApp=>_Inventor;
        
        /// <summary>
        /// inventor's <see cref="Document"/> object
        /// </summary>
        public Document ActiveDocument => _Inventor.ActiveDocument;

        /// <summary>
        /// create a new part document
        /// </summary>
        /// <returns></returns>
        public PartDocument NewPart(string templateFileName="",bool CreateVisible=true)
        {
            Documents docs = InventorApp.Documents;
            return (PartDocument)docs.Add(DocumentTypeEnum.kPartDocumentObject,templateFileName,CreateVisible);
        }

        /// <summary>
        /// create a new assembly
        /// </summary>
        /// <param name="templateFileName"></param>
        /// <param name="CreateVisible"></param>
        /// <returns></returns>
        public AssemblyDocument NewAssembly(string templateFileName = "", bool CreateVisible = true)
        {
            Documents docs = InventorApp.Documents;
            return (AssemblyDocument)docs.Add(DocumentTypeEnum.kAssemblyDocumentObject, templateFileName, CreateVisible);
        }

        /// <summary>
        /// create a new drawing. 
        /// </summary>
        /// <param name="templateFileName">if left as default value application will honor the user settings to create idw or dwg files</param>
        /// <param name="CreateVisible"></param>
        /// <returns></returns>
        public DrawingDocument NewDrawing(string templateFileName = "", bool CreateVisible = true)
        {
            Documents docs = InventorApp.Documents;
            return (DrawingDocument)docs.Add(DocumentTypeEnum.kDrawingDocumentObject, templateFileName, CreateVisible);
        }
    }
}
