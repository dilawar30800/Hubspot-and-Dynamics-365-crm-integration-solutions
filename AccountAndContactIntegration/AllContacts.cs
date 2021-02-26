using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace ContactIntegration
{

    public class HubspotOwnerId
    {
        public double value { get; set; }
        
    }

    public class Firstname
    {
        public string value { get; set; }
    }

    public class Lastmodifieddate
    {
        public double value { get; set; }
    }

    public class Createdate
    {
        public double value { get; set; }
    }

    public class Lifecyclestage
    {
        public string value { get; set; }
    }

    public class Email
    {
        public string value { get; set; }
    }

    public class Lastname
    {
        public string value { get; set; }
    }

    public class OriginatingSource
    {
        public string value { get; set; }
    }

    public class Jobtitle
    {
        public string value { get; set; }
    }

    public class Phone
    {
        public string value { get; set; }
    }

    public class LeadSource
    {
        public string value { get; set; }
    }

    public class Company
    {
        public string value { get; set; }
    }

    public class Country
    {
        public string value { get; set; }
    }

    public class City
    {
        public string value { get; set; }
    }

    public class ImportName
    {
        public string value { get; set; }
    }

    public class IcpPrefixTitle
    {
        public string value { get; set; }
    }

    public class Ext
    {
        public string value { get; set; }
    }

    public class LifecycleStatus
    {
        public string value { get; set; }
    }
   
    public class Properties
    {
        public Firstname firstname { get; set; }
        public Lastmodifieddate lastmodifieddate { get; set; }
        public Createdate createdate { get; set; }
        public Lifecyclestage lifecyclestage { get; set; }
        public Email email { get; set; }
        public Lastname lastname { get; set; }
        public OriginatingSource originating_source { get; set; }
        public Jobtitle jobtitle { get; set; }
        public Phone phone { get; set; }
        public LeadSource lead_source { get; set; }
        public Company company { get; set; }
        public Country country { get; set; }
        public City city { get; set; }
        public ImportName import_name { get; set; }
        public IcpPrefixTitle icp_prefix_title { get; set; }
        public Ext ext_ { get; set; }
        public LifecycleStatus lifecycle_status { get; set; }
        public HubspotOwnerId hubspot_owner_id { get; set; }
    }
   
    public class contact
    {
        public object addedAt { get; set; }
        public int vid { get; set; }
        [DataMember(Name="canonical-vid")]
        public int CanonicalVid { get; set; }
        [DataMember(Name="portal-id")]
        public int PortalId { get; set; }
        [DataMember(Name="is-contact")]
        public bool IsContact { get; set; }
        public Properties properties { get; set; }

    }
    [DataContract]
    public class AllContacts
    {
        [DataMember(Name = "contacts")]
        public List<contact> contacts { get; set; }
        [DataMember(Name="has-more")]
        public bool HasMore { get; set; }
        [DataMember(Name="vid-offset")]
        public int VidOffset { get; set; }
        [DataMember(Name="time-offset")]
        public long TimeOffset { get; set; }
    }

    public class ContactInfo
    {
        public string firstname { get; set; }
        public double lastmodifieddate { get; set; }
        public double createdate { get; set; }
        public string lifecyclestage { get; set; }
        public string email { get; set; }
        public string lastname { get; set; }
        public string originating_source { get; set; }
        public string jobtitle { get; set; }
        public string phone { get; set; }
        public string lead_source { get; set; }
        public string company { get; set; }
        public string country { get; set; }
        public string city { get; set; }
        public string import_name { get; set; }
        public string icp_prefix_title { get; set; }
        public string ext_ { get; set; }
        public string lifecycle_status { get; set; }
        public double company_id { get; set; }
        public Guid owner_guid { get; set; }
        public Guid contact_guid { get; set; }
        public bool is_exsiting { get; set; }
        public Guid Account_guid { get; set; }

    }


    public class ContactAssociation
    {
        public List<double> results { get; set; }

    }

}
