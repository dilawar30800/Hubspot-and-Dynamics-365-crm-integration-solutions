using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace AccountAndContactIntegration
{
    [DataContract]
  public  class Owner
    {

        [DataMember]
        public int ownerId { get; set; }
        [DataMember]
        public string type { get; set; }
        [DataMember]
        public string firstName { get; set; }
        [DataMember]
        public string lastName { get; set; }
        [DataMember]
        public string email { get; set; }
   
    }


}
