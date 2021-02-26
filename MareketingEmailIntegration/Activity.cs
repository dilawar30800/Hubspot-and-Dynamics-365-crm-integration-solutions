using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HubspotDynamics365CustomIntegration1
{
    class Activity : IComparable<Activity>
    {
        public Guid activityid { get; set; }
        public double createdon { get; set; }
        public string existingDetails { get; set; }

        public string existingSubject { get; set; }
        public string id { get; set; }
        public bool isExistInDynamics { get; set; }
        public List<Events> processing { get; set; }
        public List<Events> tracking { get; set; }

        public double emailCampaignId { get; set; }
        public double emailCampaignGroupId { get; set; }

        public string subject { get; set; }
        public string recipient { get; set; }
        public Guid contact_guid { get; set; }
        public int CompareTo(Activity other)
        {
            if (this.createdon < other.createdon)
            {
                return -1;
            }
            else if (this.createdon > other.createdon)
            {
                return 1;
            }
            else
            {
                return 0;
            }
        }

    }
    public class ContactInfo
    {
        public string email;
        public Guid contacti_id;
    }
}
