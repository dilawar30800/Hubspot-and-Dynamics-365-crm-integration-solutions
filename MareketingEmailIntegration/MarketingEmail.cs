using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace HubspotDynamics365CustomIntegration1
{
    [DataContract]
    public class SentBy
    {
        [DataMember]
        public string id { get; set; }
        [DataMember]
        public double created { get; set; }

    }
    [DataContract]
    public class ObsoletedBy
    {
        [DataMember]
        public string id { get; set; }
        [DataMember]
        public object created { get; set; }

    }
    [DataContract]
    public class CausedBy
    {
        [DataMember]
        public string id { get; set; }
        [DataMember]
        public long created { get; set; }

    }
    [DataContract]
    public class Event
    {
        [DataMember]
        public string appName { get; set; }
        [DataMember]
        public string deviceType { get; set; }
        [DataMember]
        public int? duration { get; set; }
        [DataMember]
        public string id { get; set; }
        [DataMember]
        public string userAgent { get; set; }
        [DataMember]
        public double created { get; set; }
        [DataMember]
        public string recipient { get; set; }
        [DataMember]
        public object smtpId { get; set; }
        [DataMember]
        public SentBy sentBy { get; set; }
        [DataMember]
        public string type { get; set; }

        [DataMember]
        public string category { get; set; }


        [DataMember]
        public int portalId { get; set; }
        [DataMember]
        public bool filteredEvent { get; set; }
        [DataMember]
        public int appId { get; set; }
        [DataMember]
        public int emailCampaignId { get; set; }
        [DataMember]
        public int emailCampaignGroupId { get; set; }
        [DataMember]
        public int? attempt { get; set; }
        [DataMember]
        public string response { get; set; }
        [DataMember]
        public string referer { get; set; }
        [DataMember]
        public object linkId { get; set; }
        [DataMember]
        public string url { get; set; }
        [DataMember]
        public List<subscriptions> subscriptions { get; set; }
        [DataMember]
        public string portalSubscriptionStatus { get; set; }
        [DataMember]
        public string source { get; set; }
        [DataMember]
        public string status { get; set; }
        [DataMember]
        public ObsoletedBy obsoletedBy { get; set; }
        [DataMember]
        public List<string> replyTo { get; set; }
        [DataMember]
        public string from { get; set; }
        [DataMember]
        public string subject { get; set; }
        [DataMember]
        public List<object> cc { get; set; }
        [DataMember]
        public List<object> bcc { get; set; }
        [DataMember]
        public CausedBy causedBy { get; set; }
        [DataMember]
        public string dropMessage { get; set; }
        [DataMember]
        public string dropReason { get; set; }
    }
    [DataContract]
    public class MarketingEmail
    {
        [DataMember]
        public bool hasMore { get; set; }
        [DataMember]
        public string offset { get; set; }
        [DataMember]
        public List<Event> events { get; set; }

    }

    [DataContract]
    public partial class subscriptions
    {
        [DataMember]
        public long id { get; set; }

        [DataMember]
        public string status { get; set; }

        [DataMember]
        public legalBasisChange legalBasisChange { get; set; }
    }
    [DataContract]
    public partial class legalBasisChange
    {
        [DataMember]
        public string legalBasisType { get; set; }

        [DataMember]
        public string legalBasisExplanation { get; set; }

        [DataMember]
        public string optState { get; set; }
    }
}
