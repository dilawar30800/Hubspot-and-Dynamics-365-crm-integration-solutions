using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace AccountAndContactIntegration
{
    [DataContract]
    public class Country
    {
        [DataMember]
        public string value { get; set; }
    }
    [DataContract]
    public class City
    {
        [DataMember]
        public string value { get; set; }
    }
    [DataContract]
    public class Createdate
    {
        [DataMember]
        public double value { get; set; }
    }
    [DataContract]
    public class Type
    {
        [DataMember]
        public string value { get; set; }
    }
    public class State
    {
        [DataMember]
        public string value { get; set; }
    }
    [DataContract]
    public class Website
    {
        public string value { get; set; }
    }


    [DataContract]
    public class Phone
    {
        [DataMember]
        public string value { get; set; }
    }
    [DataContract]
    public class Domain
    {
        [DataMember]
        public string value { get; set; }
    }
    [DataContract]
    public class Name
    {
        [DataMember]
        public string value { get; set; }
    }
    [DataContract]
    public class MigratedFrom
    {
        [DataMember]
        public string value { get; set; }
    }

    [DataContract]
    public class Properties
    {
        [DataMember]
        public Country country { get; set; }
        [DataMember]
        public City city { get; set; }
        [DataMember]
        public Createdate createdate { get; set; }
        [DataMember]
        public Type type { get; set; }
        [DataMember]
        public State state { get; set; }
        [DataMember]
        public Website website { get; set; }
        [DataMember]
        public MigratedFrom migrated_from { get; set; }
        [DataMember]
        public Phone phone { get; set; }
        [DataMember]
        public Domain domain { get; set; }
        [DataMember]
        public Name name { get; set; }
    }
    [DataContract]
    public class Company
    {
        [DataMember]
        public int portalId { get; set; }
        [DataMember]
        public long companyId { get; set; }
        [DataMember]
        public bool isDeleted { get; set; }
        [DataMember]
        public Properties properties { get; set; }
    }
    [DataContract]
    public class Companies
    {
        [DataMember]
        public List<Company> companies { get; set; }
        [DataMember]
        public bool hasMore { get; set; }
        [DataMember]
        public int offset { get; set; }
        [DataMember]
        public int total { get; set; }
    }


}
