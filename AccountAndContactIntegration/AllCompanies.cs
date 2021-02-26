using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccountAndContactIntegration
{
    public class HubspotOwnerId
    {
        public double value { get; set; }
        public object timestamp { get; set; }
        public string source { get; set; }
        public string sourceId { get; set; }
        public int updatedByUserId { get; set; }
    }
    public class Country
    {
        public string value { get; set; }
    }

    public class City
    {
        public string value { get; set; }
    }
    public class Createdate
    {
        public double value { get; set; }
    }

    public class Description
    {
        public string value { get; set; }
    }
    public class State
    {
        public string value { get; set; }
    }
    public class Zip
    {
        public string value { get; set; }
    }
    public class Website
    {
        public string value { get; set; }
    }

    public class Address
    {
        public string value { get; set; }
    }
    public class Type
    {
        public string value { get; set; }
    }
    public class MigratedFrom
    {
        public string value { get; set; }
    }
    public class Domain
    {
        public string value { get; set; }
    }

    public class Name
    {
        public string value { get; set; }
    }

    public class Phone
    {
        public string value { get; set; }
    }

    public class Industry
    {
        public string value { get; set; }

    }

    public class Address2
    {
        public string value { get; set; }
    }


    public class Properties
    {
        public Country country { get; set; }
        public City city { get; set; }
        public Createdate createdate { get; set; }
        public Description description { get; set; }
        public State state { get; set; }
        public Zip zip { get; set; }
        public Website website { get; set; }
        public Address address { get; set; }
        public Domain domain { get; set; }
        public Name name { get; set; }
        public Phone phone { get; set; }
        public Industry industry { get; set; }
        public Address2 address2 { get; set; }
        public Type type { get; set; }
        public MigratedFrom migrated_from { get; set; }
        public HubspotOwnerId hubspot_owner_id { get; set; }
    }

    public class Result
    {
        public int portalId { get; set; }
        public object companyId { get; set; }
        public bool isDeleted { get; set; }
        public Properties properties { get; set; }
        
    }

    public class AllCompanies
    {
        public List<Result> results { get; set; }
        public bool hasMore { get; set; }
        public int offset { get; set; }
        public int total { get; set; }
    }
    public class Account
    {
        public string owner_email { get; set; }
        public string country { get; set; }
        public string city { get; set; }
        public double createdate { get; set; }
        public string description { get; set; }
        public string state { get; set; }
        public string zip { get; set; }
        public string website { get; set; }
        public string address { get; set; }
        public string domain { get; set; }
        public string name { get; set; }
        public string phone { get; set; }
        public string industry { get; set; }
        public string address2 { get; set; }
        public string type { get; set; }
        public string migrated_from { get; set; }
        public Guid dynamics_owner_guid { get; set; }
        public Guid dynamics_account_guid { get; set; }
        public Boolean is_exist { get; set; }
    }

    public class AccountInfo
    {
        public string website;
        public Guid account_id;
    }
}
