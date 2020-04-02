using Inventor;
using System.Collections;
using System.Linq;

namespace InventorToolBox
{
    public interface IBomManager { }
    public static class ExtensionIBomManger
    {
        public static int GetBomQty(this IBomManager bomManager,AssemblyDocument assembly,Document targetDoc, BOMViewTypeEnum bomViewType)
        {
            //Get Bom Object
            IEnumerable bom= (IEnumerable)assembly.ComponentDefinition.BOM.BOMViews["Model Data"];
            switch (bomViewType)
            {
                case BOMViewTypeEnum.kModelDataBOMViewType:
                    break;
                case BOMViewTypeEnum.kStructuredBOMViewType: bom = (IEnumerable)assembly.ComponentDefinition.BOM.BOMViews["Structured"];
                    break;
                case BOMViewTypeEnum.kPartsOnlyBOMViewType: bom = (IEnumerable)assembly.ComponentDefinition.BOM.BOMViews["Parts Only"];
                    break;
                default:
                    break;
            }

            foreach (BOMRow row in bom)
            {
                if (targetDoc.InternalName==row.ReferencedFileDescriptor.ReferencedFileInternalName)
                {
                    return 2;

                }
            }
            return 1;
        }
    }
}
