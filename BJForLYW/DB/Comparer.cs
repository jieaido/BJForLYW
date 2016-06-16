using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BJForLYW.DB
{
    class PartComparer: IEqualityComparer<Part>

    {
        public bool Equals(Part x, Part y)
        {
            if (x.Partid==y.Partid)
            {
                return true;
            }
            return false;
        }

        public int GetHashCode(Part obj)
        {
            return obj.ToString().Length;
        }
    }
}
