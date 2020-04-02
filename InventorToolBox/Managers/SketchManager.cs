using Inventor;

namespace InventorToolBox
{
    public class SketchManager:IManager
    {
        public Sketch NewSketch(PartDocument part, KMainPlane plane, bool ProjectEdges = false)
        {
            // get the component definition
            PartComponentDefinition partDef = part.ComponentDefinition;

            //get the workplane
            WorkPlane workPlane = partDef.WorkPlanes[plane];

            //create a new sketch
            return partDef.Sketches.Add(workPlane, ProjectEdges) as Sketch;
        }
    }
}
