﻿using Inventor;
using System;
using System.Collections.Generic;

namespace InventorToolBox

{
    public static class AssemblyDocumentExtensions
    {
        #region private fields

        private static List<ComponentOccurrence> _list = new List<ComponentOccurrence>();
        #endregion

        #region private methode/fuctions

        /// <summary>
        /// recursively processes a document and adds componets to a private filed <see cref="_list"/>
        /// </summary>
        /// <param name="componentOccurrences"></param>
        /// <param name="targetDoc"></param>
        private static void CalculateAllNonPhantomNonReferencedOccurances(ComponentOccurrences componentOccurrences, object targetDoc)
        {
            foreach (ComponentOccurrence occurrence in componentOccurrences)
            {
                if (occurrence.Definition.BOMStructure != BOMStructureEnum.kReferenceBOMStructure
                   &&
                   occurrence.Definition.BOMStructure != BOMStructureEnum.kPhantomBOMStructure)
                {
                    if (occurrence.Definition.Document == targetDoc)
                    {
                        _list.Add(occurrence);
                    }
                    else if (occurrence.DefinitionDocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                    {
                        CalculateAllNonPhantomNonReferencedOccurances(occurrence.Definition.Occurrences, targetDoc);
                    }
                }
            }
        }
        #endregion

        /// <summary>
        /// returns BOMQuantity of specified document in the assembly document.
        /// BOMQuantity gives you access to quantity units and such.
        /// </summary>
        /// <param name="targetDoc">Target document to be queried should be non-phantom, non-reference</param>
        /// <param name="assembly">assembly where the document resides</param>
        /// <param name="bomViewType">type of bom view defined in inventor assembly environment</param>
        /// <returns>BOMQuantity</returns>
        public static BOMQuantity GetBomQuantity(this AssemblyDocument assembly, Document targetDoc , BOMViewTypeEnum bomViewType)
        {
            if (targetDoc == null)
                throw new ArgumentNullException(nameof(targetDoc), "Null argument");
            if (assembly == null)
                throw new ArgumentNullException(nameof(assembly), "Null argument");

            if (targetDoc.DocumentType == DocumentTypeEnum.kUnknownDocumentObject)
                throw new ArgumentException(nameof(targetDoc), "Unknown document");

            //Get the ActiveLevelOfDetailRepresentation Name
            string MyLOD_Name;
            MyLOD_Name = assembly.ComponentDefinition.RepresentationsManager.ActiveLevelOfDetailRepresentation.Name;

            if (MyLOD_Name != "Master")
                //activate master because only it can do the trick
                assembly.ComponentDefinition.RepresentationsManager.LevelOfDetailRepresentations[1].Activate();

            //Get Bom Object
            var oBOM = assembly.ComponentDefinition.BOM;
            //define a bomView object
            BOMView bomView = assembly.ComponentDefinition.BOM.BOMViews["Structured"];
            switch (bomViewType)
            {
                case BOMViewTypeEnum.kModelDataBOMViewType:
                    break;
                case BOMViewTypeEnum.kStructuredBOMViewType:
                    //Make sure structured view is enabled
                    oBOM.StructuredViewEnabled = true;
                    oBOM.StructuredViewFirstLevelOnly = true;
                    bomView = assembly.ComponentDefinition.BOM.BOMViews["Structured"];
                    break;
                case BOMViewTypeEnum.kPartsOnlyBOMViewType:
                    //Make sure parts only view is enabled
                    oBOM.PartsOnlyViewEnabled = true;
                    bomView = assembly.ComponentDefinition.BOM.BOMViews["Parts Only"];
                    break;
                default:
                    break;
            }
            //look for the targetDoc in assembly
            foreach (BOMRow row in bomView.BOMRows)
            {
                if (row.ComponentDefinitions[1].Document == targetDoc)
                    return row.ComponentDefinitions[1].BOMQuantity;
            }
            //at this stage the targetDoc is not found int the assembly
            return null;
        }

        /// <summary>
        /// get number of a document in a assembly but dont look into sub-assemblies
        /// </summary>
        /// <param name="targetDoc"></param>
        /// <param name="assembly"></param>
        /// <param name="countPhantomAndReference">if document is set to phantom or reference</param>
        /// <returns></returns>
        public static int GetStructuredQuantity(this AssemblyDocument assembly, Document targetDoc,  bool countPhantomAndReference = false)
        {
            if (targetDoc == null)
                throw new ArgumentNullException(nameof(targetDoc), "Null argument");
            if (assembly == null)
                throw new ArgumentNullException(nameof(assembly), "Null argument");

            if (targetDoc.DocumentType == DocumentTypeEnum.kUnknownDocumentObject)
                return 0;
            int counter = 0;

            if (countPhantomAndReference)
            {
                foreach (ComponentOccurrence occurrence in assembly.ComponentDefinition.Occurrences.AllReferencedOccurrences[targetDoc])
                {
                    counter++;
                }
            }
            else
            {
                foreach (ComponentOccurrence occurrence in assembly.AllNonPhantomNonReferencedOccurances(targetDoc))
                {
                    counter++;
                }
            }
            return counter;
        }

        /// <summary>
        /// get number of a document in a assembly,recursively looks into sub-assemblies
        /// </summary>
        /// <param name="targetDoc"></param>
        /// <param name="assembly"></param>
        /// <returns>int32 number of targetDoc in the assembly</returns>
        public static int GetPartsOnlyQuantity(this AssemblyDocument assembly, Document targetDoc, bool countPhantomAndReference = false)
        {
            if (targetDoc == null)
                throw new ArgumentNullException(nameof(targetDoc), "Null argument");
            if (assembly == null)
                throw new ArgumentNullException(nameof(assembly), "Null argument");

            if (targetDoc.DocumentType == DocumentTypeEnum.kUnknownDocumentObject)
                return 0;
            int counter = 0;

            if (countPhantomAndReference)
            {
                foreach (ComponentOccurrence occurrence in assembly.ComponentDefinition.Occurrences.AllReferencedOccurrences[targetDoc])
                {
                    counter++;
                }
            }
            else
            {
                foreach (ComponentOccurrence occurrence in assembly.AllNonPhantomNonReferencedOccurances(targetDoc))
                {
                    counter++;
                }
            }
            return counter;
        }

        /// <summary>
        /// list of occuraces that are not phantom nor are set as referenced in their document settings.
        /// </summary>
        /// <param name="targetDoc">the document that needs to be searched for</param>
        /// <returns>List<ComponentOccurrence></returns>
        public static List<ComponentOccurrence> AllNonPhantomNonReferencedOccurances(this AssemblyDocument assembly, object targetDoc)
        {
            ComponentOccurrences componentOccurrences = assembly.ComponentDefinition.Occurrences;
            CalculateAllNonPhantomNonReferencedOccurances(componentOccurrences, targetDoc);
            return _list;
        }
    }
}