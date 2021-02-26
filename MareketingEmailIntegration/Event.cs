using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HubspotDynamics365CustomIntegration1
{
    class Events
    {
        public string eventId { get; set; }
        public string type { get; set; }
        public string category { get; set; }
        public double createdon { get; set; }
        public string dropMessage { get; set; }
        public string dropReason { get; set; }
        public string portalSubscriptionStatus { get; set; }
        public List<subscriptions> Subscriptions { get; set; }


    }
}
