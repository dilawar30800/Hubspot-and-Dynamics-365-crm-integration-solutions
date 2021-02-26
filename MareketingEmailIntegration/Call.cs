using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace HubspotDynamics365CustomIntegration1
{

    public class Engagement
    {
        public bool active { get; set; }
        public int ownerId { get; set; }
        public string type { get; set; }
        public long timestamp { get; set; }
    }

    public class Associations
    {
        public List<int> contactIds { get; set; }
       
        public List<double> companyIds { get; set; }
      
    }


    public class Metadata
    {
       
        public string body { get; set; }
    }

   
    public class Call
    {

        public Engagement engagement { get; set; }
       
        public Associations associations { get; set; }
        
        public Metadata metadata { get; set; }
    }
}
