using Inventor;
using System.Collections.Generic;

namespace InventorToolBox.Managers
{
    /// <summary>
    /// Extensions on invnetor Interfaces
    /// </summary>
    public static class ComponentOccurancesExtension
    {
        private static List<ComponentOccurrence> _list = new List<ComponentOccurrence>();

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
        /// <summary>
        /// list of occuraces that are not phantom or referenced in their document settings.
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
