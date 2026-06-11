using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;

namespace Convert2DTo3D.Utils
{
    public class ElementEqualityComparer : IEqualityComparer<Element>
    {
        public bool Equals(Element x, Element y)
        {
            if (x != null && y != null)
                return x.UniqueId.Equals(y.UniqueId, StringComparison.Ordinal);
            return false;
        }

        public int GetHashCode(object obj)
        {
            return obj.GetHashCode();
        }

        public int GetHashCode(Element obj)
        {
            return obj.GetHashCode();
        }
    }
}