using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HubspotDynamics365CustomIntegration1
{
    class PhoneCall
    {
        public Guid guid { get; set; }
        public string type { get; set; }
        public DateTime createdOn { get; set; }
        public DateTime scheduledEnd { get; set; }
        public DateTime scheduleStart { get; set; }
        public string subject { get; set; }
        public string description { get; set; }
        public Entity from { get; set; }
        public Entity to { get; set; }
        public EntityReference regarding { get; set; }
        public EntityReference owner { get; set; }
        public string statecode { get; set; }
        public Guid contact_id { get; set; }
        public Guid owner_id { get; set; }
        public string contact_email { get; set; }
        public string owner_email { get; set; }
        public bool isfromhubspot { get; set;}

        public decimal new_hubspotactivityid { get; set; }

        public void show(PhoneCall activity)
        {
            //Console.Clear();
            Console.WriteLine("\nCreateOn " + activity.createdOn + "\nScheduledEnd " + activity.scheduledEnd + "\nSubject " + activity.subject + "\nDescription "+activity.description+"\n From "+activity.from.ToString()+"\nTo "+activity.to.ToString());
        }
        

    
    }
}
