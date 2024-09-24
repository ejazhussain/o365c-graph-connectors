using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace O365C.GraphConnector.MSLearn.Helpers
{
    public static class ModuleHelper
    {
        public static object GetStarRating(int count)
        {
            switch (count)
            {
                case 0:
                    return "";
                case 1:
                    return "⭐️";
                case 2:
                    return "⭐️⭐️";
                case 3:
                    return "⭐️⭐️⭐️";
                case 4:
                    return "⭐️⭐️⭐️⭐️ ";
                case 5:
                    return "⭐️⭐️⭐️⭐️⭐️";
                default:
                    return "";
            }
        }
    }
}
