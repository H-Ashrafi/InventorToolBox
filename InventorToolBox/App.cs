using Inventor;
using System;
namespace InventorToolBox
{
    public class App
    {
        #region private fields

        private readonly static object _lock = new object();
        private static Application _Inventor;
        private static App _Instance;
        #endregion

        #region private methods
        private void CheckFileName(string fullFileName, string extensin)
        {
            if (string.IsNullOrWhiteSpace(fullFileName))
                throw new ArgumentNullException(nameof(fullFileName), "string was null or empty");
            if (System.IO.File.Exists(fullFileName))
                throw new InvalidOperationException($"{fullFileName} already exists", new ArgumentException("file already exists on the address provided", nameof(fullFileName)));
            if (System.IO.Path.GetExtension(fullFileName).ToUpper() != extensin.ToUpper())
                throw new ArgumentException("extension of the file provided is wrong", nameof(fullFileName));
        }

        private void SetApplication()
        {
            try
            {
                try
                {
                    _Inventor = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
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
        #endregion

        #region constructor
        public static App Start()
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
        #endregion

        #region pulbic properties


        /// <summary>
        /// inventor's <see cref="Application"/> object
        /// </summary>
        public Application Inventor => _Inventor;

        /// <summary>
        /// inventor's <see cref="Document"/> object
        /// </summary>
        public Document ActiveDocument => _Inventor.ActiveDocument;
        #endregion

        #region public methods
        /// <summary>
        /// create a new part document
        /// </summary>
        /// <returns></returns>
        public PartDocument NewPart(string fullFileName, string templateFileName = "", bool CreateVisible = true)
        {
            CheckFileName(fullFileName, ".ipt");
            Documents docs = Inventor.Documents;
            var part = (PartDocument)docs.Add(DocumentTypeEnum.kPartDocumentObject, templateFileName, CreateVisible);
            part.SaveAs(fullFileName, false);
            return part;
        }

        /// <summary>
        /// create a new assembly
        /// </summary>
        /// <param name="templateFileName"></param>
        /// <param name="CreateVisible"></param>
        /// <returns></returns>
        public AssemblyDocument NewAssembly(string fullFileName, string templateFileName = "", bool CreateVisible = true)
        {
            CheckFileName(fullFileName, ".iam");
            Documents docs = Inventor.Documents;
            var assy = (AssemblyDocument)docs.Add(DocumentTypeEnum.kAssemblyDocumentObject, templateFileName, CreateVisible);
            assy.SaveAs(fullFileName, false);
            return assy;
        }

        /// <summary>
        /// create a new drawing. 
        /// </summary>
        /// <param name="templateFileName">if left as default value application will honor the user settings to create idw or dwg files</param>
        /// <param name="CreateVisible"></param>
        /// <returns></returns>
        public DrawingDocument NewDrawing(string fullFileName, string templateFileName = "", bool CreateVisible = true)
        {
            CheckFileName(fullFileName, ".idw");
            Documents docs = Inventor.Documents;
            return (DrawingDocument)docs.Add(DocumentTypeEnum.kDrawingDocumentObject, templateFileName, CreateVisible);
        }

        /// <summary>
        /// Open an existing document from disk.
        /// </summary>
        /// <param name="fullFileName"></param>
        /// <param name="visible"></param>
        /// <returns></returns>
        public Document Open(string fullFileName, bool visible = true)
        {
            if (!System.IO.File.Exists(fullFileName))
                throw new InvalidOperationException($"{fullFileName} does not exist");
            
            //Open an existing document from disk.
            var oDoc = Inventor.Documents.Open(fullFileName, visible);
            return oDoc;
        }
        #endregion
    }
}
