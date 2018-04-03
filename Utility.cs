using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace NavisClash
{
    class Utility
    {
        internal static string GetConnectionString()
        {
            string returnValue = null;

            returnValue = "Data Source = (localdb)\\ProjectsV13; Initial Catalog = ClashDB; Integrated Security = True; Connect Timeout = 30; Encrypt = False; TrustServerCertificate = True; ApplicationIntent = ReadWrite; MultiSubnetFailover = False";
            return returnValue;
        }
    }
}
