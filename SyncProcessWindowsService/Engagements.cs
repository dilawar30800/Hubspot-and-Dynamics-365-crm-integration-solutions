using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace SyncProcessWindowsService
{
    [DataContract]
    public class Engagement
    {
        [DataMember]
        public decimal id { get; set; }
        [DataMember]
        public bool active { get; set; }
        [DataMember]
        public object createdAt { get; set; }
        [DataMember]
        public object lastUpdated { get; set; }
        [DataMember]
        public string type { get; set; }
        [DataMember]
        public object timestamp { get; set; }
        [DataMember]
        public string bodyPreview { get; set; }
        [DataMember]
        public string bodyPreviewHtml { get; set; }
        [DataMember]
        public int ownerId { get; set; }
        [DataMember]
        public string sourceId { get; set; }

    }

    [DataContract]
    public class Associations
    {
        [DataMember]
        public List<int> contactIds { get; set; }
    }
    [DataContract]
    public class ScheduledTask
    {
        [DataMember]
        public long engagementId { get; set; }
        [DataMember]
        public int portalId { get; set; }
        [DataMember]
        public string engagementType { get; set; }
        [DataMember]
        public string taskType { get; set; }
        [DataMember]
        public long timestamp { get; set; }
        [DataMember]
        public string uuid { get; set; }
    }
    [DataContract]
    public class Metadata
    {
        [DataMember]
        public string body { get; set; }
        [DataMember]
        public string status { get; set; }
        [DataMember]
        public string subject { get; set; }
        [DataMember]
        public string taskType { get; set; }
        [DataMember]
        public List<long> reminders { get; set; }
        [DataMember]
        public bool sendDefaultReminder { get; set; }
        [DataMember]
        public string priority { get; set; }
        [DataMember]
        public bool isAllDay { get; set; }
        [DataMember]
        public string title { get; set; }
        [DataMember]
        public long startTime { get; set; }
        [DataMember]
        public long endTime { get; set; }
        [DataMember]
        public string disposition { get; set; }
    }
    [DataContract]
    public class Result
    {
        [DataMember]
        public Engagement engagement { get; set; }
        [DataMember]
        public Associations associations { get; set; }
        [DataMember]
        public Metadata metadata { get; set; }
    }
    [DataContract]
    public class Engagements
    {
        [DataMember]
        public List<Result> results { get; set; }
        [DataMember]
        public bool hasMore { get; set; }
        [DataMember]
        public int offset { get; set; }
        [DataMember]
        public int total { get; set; }
    }
}
