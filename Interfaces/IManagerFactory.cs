using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InventorToolBox.Interfaces
{
    //abstract factory
    public abstract class IManagerFactory
    {
        public abstract ISketchManager CreatSketchManager();
        public abstract IDocumentManager CreatDocumentManager();
        public abstract IAssemblyManager CreatAssemblyManager();
        public abstract IPartManager CreatPartManager();
        public abstract IPropertyManager CreatPropertyManager();
    }

    //abstract mangers (products)
    public interface ISketchManager { }
    public interface IDocumentManager { }
    public interface IAssemblyManager { }
    public interface IPartManager { }
    public interface IPropertyManager { }
}
