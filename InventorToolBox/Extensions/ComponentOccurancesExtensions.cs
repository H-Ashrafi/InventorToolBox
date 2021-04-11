using Inventor;
using System.Collections.Generic;

namespace InventorToolBox
{
    /// <summary>
    /// Extensions on <see cref="ComponentOccurrences"/>
    /// </summary>
    public static class ComponentOccurancesExtension
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
        /// list of occuraces that are not phantom nor are set as referenced in their document settings.
        /// </summary>
        /// <param name="targetDoc">the document that needs to be searched for</param>
        /// <returns>List<ComponentOccurrence></returns>
        public static List<ComponentOccurrence> AllNonPhantomNonReferencedOccurances(this ComponentOccurrences componentOccurrences, object targetDoc)
        {
            CalculateAllNonPhantomNonReferencedOccurances(componentOccurrences, targetDoc);
            return _list;
        }
    }
}
