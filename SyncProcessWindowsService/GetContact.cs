using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace SyncProcessWindowsService
{
    [DataContract]
    public class Associatedcompanyid
    {
        [DataMember]
        public string value { get; set; }
    }
    [DataContract]
    public class Properties1
    {
        [DataMember]
        public Associatedcompanyid associatedcompanyid { get; set; }
    }
    [DataContract]
    public class GetContact
    {
        [DataMember]
        public int vid { get; set; }

        [DataMember]
        public Properties1 properties { get; set; }
    }
}
