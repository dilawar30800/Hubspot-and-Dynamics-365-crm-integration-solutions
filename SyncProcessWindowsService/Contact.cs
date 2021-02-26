using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace SyncProcessWindowsService
{
    [DataContract]
    public class HubspotOwnerId
    {
        [DataMember]
        public string value { get; set; }
    }

    [DataContract]
    public class Lastname
    {
        [DataMember]
        public string value { get; set; }
    }

    [DataContract]
    public class Firstname
    {
        [DataMember]
        public string value { get; set; }
    }
    [DataContract]
    public class Email
    {
        [DataMember]
        public string value { get; set; }
    }


    [DataContract]
    public class Properties
    {
        [DataMember]
        public HubspotOwnerId hubspot_owner_id { get; set; }
        [DataMember]
        public Lastname lastname { get; set; }
        [DataMember]
        public Firstname firstname { get; set; }
        [DataMember]
        public Email email { get; set; }

    }
    [DataContract]
    class Contact
    {
        [DataMember]
        public int vid { get; set; }
        [DataMember]
        public Properties properties { get; set; }
    }
}
  
