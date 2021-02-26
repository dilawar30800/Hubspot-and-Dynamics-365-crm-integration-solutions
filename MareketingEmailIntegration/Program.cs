using CustomDynamicsHubSpotIntegration;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.Serialization.Json;
using System.ServiceModel.Description;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Configuration;
using System.Collections.Specialized;

namespace HubspotDynamics365CustomIntegration1
{

    class Program
    {
        string API_Key = ConfigurationManager.AppSettings["API_Key"];
        // double time = 1590650253617; // as today 28th May 12:19 pm  
        public string FindandReplace(string Source, string Find, string Replace)
        {
            int Place = Source.IndexOf(Find);
            if (Place == -1)
            {
                return string.Empty;
            }
            else
            {

                string result = Source.Remove(Place, Find.Length).Insert(Place, Replace);
                return result;
            }
        }
        public IOrganizationService getConnection()
        {
            ClientCredentials clientCredentials1 = new ClientCredentials();
            clientCredentials1.UserName.UserName = ConfigurationManager.AppSettings["CRMUserName"];
            clientCredentials1.UserName.Password = ConfigurationManager.AppSettings["CRMPWS"];

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            OrganizationServiceProxy proxy_service = new OrganizationServiceProxy(new Uri(ConfigurationManager.AppSettings["CRMApiURL"].ToString()),
             null, clientCredentials1, null);
            var timeout = new TimeSpan(0, 10, 0);
            proxy_service.Timeout = timeout;
            IOrganizationService organizationService1 = (IOrganizationService)proxy_service;

            if (organizationService1 != null)
            {
                return organizationService1;
            }
            else
            {
                return null;
            }
        }
        public static bool isCreatedInHubspot(string description)
        {
            return description.Contains("(Activity Created From Dynamics 365)");
        }
        public List<PhoneCall> getAllEngagements(double time)
        {
            IOrganizationService service = getConnection();
            int phone_count = 0;
            int email_count = 0;
            int meeting_count = 0;
            int incoming_email_count = 0;

            List<PhoneCall> list = new List<PhoneCall>();
            int offset = 0;
            bool has_more = true;
            while (has_more)
            {
                //maximum load limit is 100 in case of recently modified  engagments 
                string url = "https://api.hubapi.com/engagements/v1/engagements/recent/modified?hapikey=" + API_Key + "&count=100&since=" + time + "&offset=" + offset;
                HttpWebRequest webrequest = (HttpWebRequest)WebRequest.Create(url);
                webrequest.Method = "GET";
                webrequest.ContentType = "application/json";
                HttpWebResponse webresponse = null;
                try
                {
                    webresponse = (HttpWebResponse)webrequest.GetResponse();
                }

                catch (Exception e)
                {
                    Console.WriteLine("Get All Engagment exception details:" + e.StackTrace);
                }
                Encoding enc = System.Text.Encoding.GetEncoding("utf-8");
                StreamReader responseStream = new StreamReader(webresponse.GetResponseStream(), enc);
                string result = string.Empty;
                result = responseStream.ReadToEnd();
                webresponse.Close();
                Engagements obj;
                using (var ms = new MemoryStream(Encoding.Unicode.GetBytes(result)))
                {
                    // Deserialization from JSON
                    DataContractJsonSerializer deserializer = new DataContractJsonSerializer(typeof(Engagements));
                    obj = (Engagements)deserializer.ReadObject(ms);

                }

                Console.WriteLine("Total Fetched Activities from HS: {0} ", obj.results.Count);

                has_more = obj.hasMore;
                offset = obj.offset;
                Console.WriteLine("Parameter has-more={0}, offset={1} ", obj.hasMore, obj.offset);
                Console.WriteLine("Preparing  Activities list");
                Console.WriteLine("Please wait it will take some time......");
                for (int i = 0; i < obj.results.Count; i++)
                {
                    //here we are getting engagements 
                    var engagementType = obj.results[i].engagement.type;
                    var engagementBody = obj.results[i].engagement.bodyPreview;
                    var engagmenetBodyHTML = obj.results[i].engagement.bodyPreviewHtml;
                    var engagementSubject = obj.results[i].metadata.subject;
                    var meetingTitle = obj.results[i].metadata.title;
                    var ownerId = obj.results[i].engagement.ownerId;
                    decimal activity_id = obj.results[i].engagement.id;

                    //checking if activity is created from dynamics plugin so to skip duplication

                    string subject;
                    if (engagementType == ActivityType.CALL.ToString())
                    {
                        subject = engagementBody;
                    }
                    else if (engagementType == ActivityType.EMAIL.ToString() || engagementType == ActivityType.INCOMING_EMAIL.ToString())
                    {
                        subject = engagementBody;
                    }
                    else if (engagementType == ActivityType.MEETING.ToString())
                    {
                        subject = engagementBody;

                    }
                    else
                    {
                        continue;
                    }


                    if (subject != null)
                    {
                        if (isCreatedInHubspot(subject))
                        { continue; }

                        else if (isActivityExistInDynamics(activity_id, engagementType))
                        {

                            continue;
                        }
                    }
                    else
                    {
                        continue;
                    }





                    DateTime engagement_createdAt = convertTimestampToDatetime((long)obj.results[i].engagement.createdAt).ToLocalTime();
                    if (ownerId == 0)
                    {
                        string log = "OwnerId is zero: acitivity id " + activity_id.ToString();
                        Console.Write(log);
                        continue;
                    }

                    string engagement_createdBy = getOwnerEmailHubSpotByOwnerId(ownerId);
                    //Console.WriteLine("Engagement Details at index {0} is \ncreated at="+engagement_createdAt+"\nBy Account Holder="+engagement_createdBy+"\nof type=" + engagementType + "\nsubject=" + engagementSubject + "\nbody=" + engagementBody + "\n", i);
                    //Console.Write("Contact Id(s):");
                    //foreach(int id in obj.results[i].associations.contactIds )
                    //{
                    //    Console.WriteLine("\n" + id );
                    //}
                    //Console.WriteLine("\n\n");

                    if (engagementType == ActivityType.CALL.ToString())
                    {
                        if (obj.results[i].associations.contactIds.Count > 0)
                        {
                            foreach (var item in obj.results[i].associations.contactIds)
                            {
                                //Console.WriteLine(ActivityType.CALL.ToString() + " you entered in condition having more than zero contact");
                                int hubspotContactId = item;
                                PhoneCall activity = new PhoneCall();
                                activity.new_hubspotactivityid = activity_id;
                                activity.type = engagementType;
                                activity.isfromhubspot = true;
                                activity.createdOn = engagement_createdAt;
                                activity.scheduledEnd = engagement_createdAt;
                                activity.subject = engagementSubject;
                                activity.description = engagementBody;
                                string email = getConactEmailHubSpotByContactId(hubspotContactId);
                                Guid contact_id = getGuidDynamicsContactByEmail(email);
                                if (contact_id.Equals(Guid.Empty))
                                {
                                    continue;
                                }
                                Guid owner_guid = getGuidDynamicsUserByEmail(engagement_createdBy);
                                if (owner_guid.Equals(Guid.Empty))
                                {
                                    continue;
                                }
                                activity.contact_email = email;
                                activity.owner_email = engagement_createdBy;
                                activity.contact_id = contact_id;
                                activity.owner_id = owner_guid;
                                //create Entity object 

                                Entity From = new Entity("activityparty");
                                From["partyid"] = new EntityReference("systemuser", owner_guid);
                                Entity To = new Entity("activityparty");
                                To["partyid"] = new EntityReference("contact", contact_id);
                                EntityReference Regarding = new EntityReference("contact", contact_id);
                                EntityReference owner_reference = new EntityReference("systemuser", owner_guid);

                                activity.to = To;
                                activity.from = From;
                                activity.regarding = Regarding;
                                activity.owner = owner_reference;
                                // Console.WriteLine("\n\n\n\n\n" + email + "\n" + contact_id + "\n" + owner_guid + "\n\n\n\n\n\n\n\n");
                                list.Add(activity);
                                phone_count++;
                            }

                        }

                    }

                    //processing Email related data
                    if (engagementType == ActivityType.EMAIL.ToString() || engagementType == ActivityType.INCOMING_EMAIL.ToString())
                    {

                        if (obj.results[i].associations.contactIds.Count > 0)
                        {
                            foreach (var item in obj.results[i].associations.contactIds)
                            {
                                //Console.WriteLine(ActivityType.EMAIL.ToString() + " you entered in condition having more than zero contact");
                                int hubspotContactId = item;
                                PhoneCall activity = new PhoneCall();
                                activity.new_hubspotactivityid = activity_id;
                                activity.type = engagementType;
                                activity.isfromhubspot = true;
                                activity.createdOn = engagement_createdAt;
                                activity.scheduledEnd = engagement_createdAt;
                                activity.subject = engagementSubject;
                                activity.description = engagementBody;
                                string email = getConactEmailHubSpotByContactId(hubspotContactId);
                                Guid contact_id = getGuidDynamicsContactByEmail(email);
                                if (contact_id.Equals(Guid.Empty))
                                {
                                    continue;
                                }
                                Guid owner_guid = getGuidDynamicsUserByEmail(engagement_createdBy);
                                if (owner_guid.Equals(Guid.Empty))
                                {
                                    continue;
                                }
                                activity.contact_email = email;
                                activity.owner_email = engagement_createdBy;
                                activity.contact_id = contact_id;
                                activity.owner_id = owner_guid;
                                //create Entity object 

                                Entity From = new Entity("activityparty");
                                From["partyid"] = new EntityReference("systemuser", owner_guid);
                                Entity To = new Entity("activityparty");
                                To["partyid"] = new EntityReference("contact", contact_id);
                                EntityReference Regarding = new EntityReference("contact", contact_id);
                                EntityReference owner_reference = new EntityReference("systemuser", owner_guid);

                                activity.to = To;
                                activity.from = From;
                                activity.regarding = Regarding;
                                activity.owner = owner_reference;
                                // Console.WriteLine("\n\n\n\n\n" + email + "\n" + contact_id + "\n" + owner_guid + "\n\n\n\n\n\n\n\n");
                                list.Add(activity);
                                if (engagementType == ActivityType.INCOMING_EMAIL.ToString())
                                {
                                    incoming_email_count++;
                                }
                                else
                                {
                                    email_count++;
                                }

                            }

                        }

                    }


                    //Meeting Capturing
                    if (engagementType == ActivityType.MEETING.ToString())
                    {
                        if (obj.results[i].associations.contactIds.Count > 0)
                        {
                            var startTime = obj.results[i].metadata.startTime;
                            var endTime = obj.results[i].metadata.endTime;
                            var dynamicSubject = obj.results[i].metadata.title;
                            foreach (var item in obj.results[i].associations.contactIds)
                            {
                                //Console.WriteLine(ActivityType.MEETING.ToString() + " you entered in condition having more than zero contact");
                                int hubspotContactId = item;
                                PhoneCall activity = new PhoneCall();
                                activity.new_hubspotactivityid = activity_id;
                                activity.type = engagementType;
                                activity.isfromhubspot = true;
                                activity.createdOn = engagement_createdAt;
                                activity.scheduledEnd = engagement_createdAt;
                                activity.subject = dynamicSubject;
                                activity.description = engagementBody;
                                string email = getConactEmailHubSpotByContactId(hubspotContactId);
                                Guid contact_id = getGuidDynamicsContactByEmail(email);
                                if (contact_id.Equals(Guid.Empty))
                                {
                                    continue;
                                }
                                Guid owner_guid = getGuidDynamicsUserByEmail(engagement_createdBy);
                                if (owner_guid.Equals(Guid.Empty))
                                {
                                    continue;
                                }
                                activity.contact_email = email;
                                activity.owner_email = engagement_createdBy;
                                activity.contact_id = contact_id;
                                activity.owner_id = owner_guid;
                                //create Entity object 

                                Entity From = new Entity("activityparty");
                                From["partyid"] = new EntityReference("systemuser", owner_guid);
                                Entity To = new Entity("activityparty");
                                To["partyid"] = new EntityReference("contact", contact_id);
                                EntityReference Regarding = new EntityReference("contact", contact_id);
                                EntityReference owner_reference = new EntityReference("systemuser", owner_guid);

                                activity.to = To;
                                activity.from = From;
                                activity.regarding = Regarding;
                                activity.owner = owner_reference;
                                // Console.WriteLine("\n\n\n\n\n" + email + "\n" + contact_id + "\n" + owner_guid + "\n\n\n\n\n\n\n\n");
                                list.Add(activity);
                                meeting_count++;
                            }

                        }

                    }

                }
            }
            Console.WriteLine("Total prepared Activities: " + list.Count);
            Console.WriteLine("Phone Call Activities are: " + phone_count);
            Console.WriteLine("Email Activities are: " + email_count);
            Console.WriteLine("Incoming Email Activities are: " + incoming_email_count);
            Console.WriteLine("Meeting Activities are: " + meeting_count);
            return list;
        }
        public Boolean isActivityExistInDynamics(decimal hubspotactivityid, string type)
        {
            ConditionExpression condition1 = new ConditionExpression();
            condition1.AttributeName = "new_hubspotactivityid";
            condition1.Operator = ConditionOperator.Equal;
            condition1.Values.Add(hubspotactivityid);

            FilterExpression filter1 = new FilterExpression();
            filter1.Conditions.Add(condition1);
            string entity = "";
            if (type == "CALL") { entity = "phonecall"; }
            else if (type == "EMAIL") { entity = "email"; }
            else if (type == "MEETING") { entity = "appointment"; }
            else { return false; }
            QueryExpression query = new QueryExpression(entity);
            query.ColumnSet.AddColumns("new_isfromhubspot");
            query.Criteria.AddFilter(filter1);
            bool result = false;
            try
            {
                IOrganizationService service = getConnection();
                EntityCollection records = service.RetrieveMultiple(query);
                if (records.Entities.Count > 0)
                {
                    result = true; ;
                }

            }
            catch (Exception e)
            {
                Console.WriteLine("Exception details in IsActivityExistinDynamics for activity id: {2} Message: {0} StackTrack {1}" + e.Message, e.StackTrace, hubspotactivityid);
            }
            return result;

        }
        public string getConactEmailHubSpotByContactId(double id)
        {
            String email = string.Empty;
            string url = "https://api.hubapi.com/contacts/v1/contact/vid/" + id + "/profile?hapikey=" + API_Key;
            HttpWebRequest webrequest = (HttpWebRequest)WebRequest.Create(url);
            webrequest.Method = "GET";
            webrequest.ContentType = "application/json";
            HttpWebResponse webresponse = null;
            try
            {
                webresponse = (HttpWebResponse)webrequest.GetResponse();
            }
            catch (Exception e)
            {
                string log = "Exception details in getConactEmailHubSpotByContactId for id : " + id.ToString();
                ErrorLogging(e, log);
                Console.WriteLine("Exception details for contact id: {2} Message: {0} StackTrack {1}" + e.Message, e.StackTrace, id);
            }
            Encoding enc = System.Text.Encoding.GetEncoding("utf-8");
            StreamReader responseStream = new StreamReader(webresponse.GetResponseStream(), enc);
            string result = string.Empty;
            result = responseStream.ReadToEnd();
            webresponse.Close();
            Contact obj;
            using (var ms = new MemoryStream(Encoding.Unicode.GetBytes(result)))
            {
                // Deserialization from JSON
                DataContractJsonSerializer deserializer = new DataContractJsonSerializer(typeof(Contact));
                obj = (Contact)deserializer.ReadObject(ms);
            }
            email = obj.properties.email.value;
            //Console.WriteLine("email of contact Id is: "+email);
            return email;
        }
        public string getOwnerEmailHubSpotByOwnerId(double id)
        {
            String email = string.Empty;
            string url = "http://api.hubapi.com/owners/v2/owners/" + id + "?hapikey=" + API_Key;
            HttpWebRequest webrequest = (HttpWebRequest)WebRequest.Create(url);
            webrequest.Method = "GET";
            webrequest.ContentType = "application/json";
            HttpWebResponse webresponse = null;
            try
            {
                webresponse = (HttpWebResponse)webrequest.GetResponse();
            }
            catch (Exception e)
            {
                string log = "Exception details in getOwnerEmailHubSpotByOwnerId for id : " + id.ToString();
                ErrorLogging(e, log);
                Console.WriteLine("Exception details for owner id: {2} Message: {0} StackTrack {1}" + e.Message, e.StackTrace, id);
            }
            Encoding enc = System.Text.Encoding.GetEncoding("utf-8");
            StreamReader responseStream = new StreamReader(webresponse.GetResponseStream(), enc);
            string result = string.Empty;
            result = responseStream.ReadToEnd();
            webresponse.Close();
            Owner obj;
            using (var ms = new MemoryStream(Encoding.Unicode.GetBytes(result)))
            {
                // Deserialization from JSON
                DataContractJsonSerializer deserializer = new DataContractJsonSerializer(typeof(Owner));
                obj = (Owner)deserializer.ReadObject(ms);

            }
            email = obj.email;
            //Console.WriteLine("email of Owner is: " + email);
            return email;
        }
        public Guid getGuidDynamicsUserByEmail(string email)
        {
            ConditionExpression condition1 = new ConditionExpression();
            condition1.AttributeName = "domainname";
            condition1.Operator = ConditionOperator.Equal;
            condition1.Values.Add(email);

            FilterExpression filter1 = new FilterExpression();
            filter1.Conditions.Add(condition1);

            QueryExpression query = new QueryExpression("systemuser");
            query.ColumnSet.AddColumns("fullname", "systemuserid", "domainname", "fullname");
            query.Criteria.AddFilter(filter1);
            Guid userId = new Guid();
            try
            {
                IOrganizationService service = getConnection();
                EntityCollection records = service.RetrieveMultiple(query);

                Entity record = records.Entities.First();
                userId = record.GetAttributeValue<Guid>("systemuserid");
                //Console.WriteLine("value at index " + userId);

            }
            catch (Exception e)
            {

                string log = "Exception details in getGuidDynamicsUserByEmail for Email : " + email;
                ErrorLogging(e, log);
                Console.WriteLine("Exception details in getGuidDynamicsUserByEmail for Email : {2} Message: {0} StackTrack {1}" + e.Message, e.StackTrace, email);
            }
            return userId;
        }
        public Guid getGuidDynamicsContactByEmail(string email)
        {
            ConditionExpression condition1 = new ConditionExpression();
            condition1.AttributeName = "emailaddress1";
            condition1.Operator = ConditionOperator.Equal;
            condition1.Values.Add(email);

            FilterExpression filter1 = new FilterExpression();
            filter1.Conditions.Add(condition1);

            QueryExpression query = new QueryExpression("contact");
            query.ColumnSet.AddColumns("emailaddress1");
            query.Criteria.AddFilter(filter1);
            Guid contactId = new Guid();
            try
            {
                IOrganizationService service = getConnection();
                EntityCollection records = service.RetrieveMultiple(query);

                if (records.Entities.Count > 0)
                {
                    Entity record = records.Entities.First();
                    contactId = record.GetAttributeValue<Guid>("contactid");
                    //Console.WriteLine("value at index " + userId);

                }

            }
            catch (Exception e)
            {
                string log = "Exception details in getGuidDynamicsContactByEmail for Email : " + email;
                ErrorLogging(e, log);
                Console.WriteLine("Exception details in getGuidDynamicsContactByEmail for Email : {2} Message: {0} StackTrack {1}" + e.Message, e.StackTrace, email);

            }
            return contactId;

        }
        //just testing in operator, shoudl be removed after testing
        public List<ContactInfo> getGuidDynamicsMultipleContactByEmail(List<string> list)
        {



            //int iterations = Convert.ToInt32(list.Count / 2000) + 1;
            List<ContactInfo> output = new List<ContactInfo>();
            

            //starting from here the batch code
            int counter = 0;
            int batchsize = 1500;
            for (int x = 0; x < Math.Ceiling((decimal)list.Count / batchsize); x++)
            {
                ConditionExpression condition1 = new ConditionExpression();
                condition1.AttributeName = "emailaddress1";
                condition1.Operator = ConditionOperator.In;
                var contacts = list.Skip(x * batchsize).Take(batchsize);
                //making request for contacts
                foreach(var contact in contacts)
                {

                    condition1.Values.Add(contact.ToString());
                }


                FilterExpression filter1 = new FilterExpression();
                filter1.Conditions.Add(condition1);

                QueryExpression query = new QueryExpression("contact");
                query.ColumnSet.AddColumns("emailaddress1");
                query.Criteria.AddFilter(filter1);
                //Guid contactId = new Guid();
                try
                {
                    IOrganizationService service = getConnection();
                    EntityCollection records = service.RetrieveMultiple(query);

                    if (records.Entities.Count > 0)
                    {
                        foreach (Entity record in records.Entities)
                        {
                            output.Add(new ContactInfo() { contacti_id = record.GetAttributeValue<Guid>("contactid"), email = record.GetAttributeValue<string>("emailaddress1") });

                        }
                        // Entity record = records.Entities.First();

                        //Console.WriteLine("value at index " + userId);

                    }

                }
                catch (Exception e)
                {
                    string log = "Exception details in getGuidDynamicsMultipleContactByEmail  : with complet list ";
                    ErrorLogging(e, log);
                    Console.WriteLine("Exception details in getGuidDynamicsContactByEmail for Email : {2} Message: {0} StackTrack {1}" + e.Message, e.StackTrace, "with this list");

                }

            }


            //for (int j = 1; j <= iterations; j++)
            //{
           


            //}
            return output;

        }
        public Activity getGuidMarketingEmailEvevtByEmailCampaignId(double emailCampaignId, string recipient)
        {

            Activity activity = new Activity();
            ConditionExpression condition = new ConditionExpression();
            condition.AttributeName = "new_emailcampaignid";
            condition.Operator = ConditionOperator.Equal;
            condition.Values.Add(Convert.ToInt32(emailCampaignId));


            FilterExpression filter = new FilterExpression();
            filter.Conditions.Add(condition);

            ConditionExpression condition1 = new ConditionExpression();
            condition1.AttributeName = "new_recipient";
            condition1.Operator = ConditionOperator.Equal;
            condition1.Values.Add(recipient);

            FilterExpression filter1 = new FilterExpression();
            filter1.Conditions.Add(condition1);

            QueryExpression query = new QueryExpression("new_marketingemail");
            query.ColumnSet.AddColumns("activityid");
            query.ColumnSet.AddColumns("new_details");
            query.ColumnSet.AddColumns("subject");
            query.Criteria.AddFilter(filter);
            query.Criteria.AddFilter(filter1);
            OrderExpression order = new OrderExpression();
            order.AttributeName = "createdon";
            order.OrderType = OrderType.Descending;
            query.Orders.Add(order);
            Guid activityId = new Guid();
            string details = string.Empty;
            activity.activityid = Guid.Empty;
            try
            {
                IOrganizationService service = getConnection();
                EntityCollection records = service.RetrieveMultiple(query);

                if (records.Entities.Count > 0)
                {
                    Entity record = records.Entities.First();
                    activityId = record.GetAttributeValue<Guid>("activityid");
                    details = record.GetAttributeValue<string>("new_details");
                    activity.activityid = activityId;
                    activity.existingDetails = details;
                    activity.existingSubject = record.GetAttributeValue<string>("subject");
                    //Console.WriteLine("value at index " + userId);

                }

            }
            catch (Exception e)
            {
                string log = "Exception details in getGUIDMarketingEmailByEmapCampaignId for Email Campaign Id : " + emailCampaignId;
                ErrorLogging(e, log);
                Console.WriteLine("Exception details in getGUIDMarketingEmailByEmapCampaignId for Email Campaign Id : {2} Message: {0} StackTrack {1}" + e.Message, e.StackTrace, emailCampaignId);

            }
            return activity;

        }
        public DateTime convertTimestampToDatetime(double timestamp)
        {
            DateTime dtDateTime = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddMilliseconds(timestamp).ToLocalTime();
            return dtDateTime;
        }
        public static void ErrorLogging(Exception ex, string param)
        {
            string strPath = @"D:\ErrorLog.txt";
            if (!File.Exists(strPath))
            {
                File.Create(strPath).Dispose();
            }
            using (StreamWriter sw = File.AppendText(strPath))
            {
                sw.WriteLine("=============Error Logging ===========");
                sw.WriteLine("===========Start============= " + DateTime.Now);
                sw.WriteLine("Details: " + param);
                sw.WriteLine("Error Message: " + ex.Message);
                sw.WriteLine("Stack Trace: " + ex.StackTrace);
                sw.WriteLine("Stack Trace: " + ex.InnerException);
                sw.WriteLine("===========End============= " + DateTime.Now);
            }
        }
        public static void CustomLogging(string param)
        {
            string strPath = @"D:\logCamp.txt";
            if (!File.Exists(strPath))
            {
                File.Create(strPath).Dispose();
            }
            using (StreamWriter sw = File.AppendText(strPath))
            {
                //sw.WriteLine("============= Logging ===========");
                //sw.WriteLine("===========Start============= " );
                sw.WriteLine("Details: " + param);
                //sw.WriteLine("===========End============= " );
            }
        }
        public static void FilteredCampaignLog(string param)
        {
            string strPath = @"D:\FilterCapaignLog.txt";
            if (!File.Exists(strPath))
            {
                File.Create(strPath).Dispose();
            }
            using (StreamWriter sw = File.AppendText(strPath))
            {
                //sw.WriteLine("============= Logging ===========");
                //sw.WriteLine("===========Start============= " );
                sw.WriteLine("Campaign Details: " + param);
                //sw.WriteLine("===========End============= " );
            }
        }
        public static void CampaignLog(string param)
        {
            string strPath = @"D:\CampaignLog.txt";
            if (!File.Exists(strPath))
            {
                File.Create(strPath).Dispose();
            }
            using (StreamWriter sw = File.AppendText(strPath))
            {
                //sw.WriteLine("============= Logging ===========");
                //sw.WriteLine("===========Start============= " );
                sw.WriteLine("Details: " + param);
                //sw.WriteLine("===========End============= " );
            }
        }
        public static void writeList(List<Activity> list)
        {
            string strPath = @"D:\CampaignsList.txt";
            if (!File.Exists(strPath))
            {
                File.Create(strPath).Dispose();
            }
            using (StreamWriter sw = File.AppendText(strPath))
            {
                foreach (var obj in list)
                {
                    DateTime dtDateTime = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddMilliseconds(obj.createdon).ToLocalTime();

                    sw.WriteLine("Created on: " + dtDateTime + " subject: " + obj.subject);
                }
                //sw.WriteLine("============= Logging ===========");
                //sw.WriteLine("===========Start============= " );

                //sw.WriteLine("===========End============= " );
            }
        }
        public static void writeEmailsBeforeFilterationLog(List<Activity> list)
        {
            string strPath = @"D:\Emails before filter.txt";
            if (!File.Exists(strPath))
            {
                File.Create(strPath).Dispose();
            }
            using (StreamWriter sw = File.AppendText(strPath))
            {
                foreach (var obj in list)
                {

                    sw.WriteLine("Email: " + obj.recipient);
                }
                //sw.WriteLine("============= Logging ===========");
                //sw.WriteLine("===========Start============= " );

                //sw.WriteLine("===========End============= " );
            }
        }
        public static void writeEmailsAfterFilterationLog(List<Activity> list)
        {
            string strPath = @"D:\Emails after filter.txt";
            if (!File.Exists(strPath))
            {
                File.Create(strPath).Dispose();
            }
            using (StreamWriter sw = File.AppendText(strPath))
            {
                foreach (var obj in list)
                {

                    sw.WriteLine("Email: " + obj.recipient);
                }
                //sw.WriteLine("============= Logging ===========");
                //sw.WriteLine("===========Start============= " );

                //sw.WriteLine("===========End============= " );
            }
        }
        public List<Activity> getMarketingEmails(double starttime, double endtime)
        {
            Activity activity = new Activity();
            // IOrganizationService service = getConnection();

            List<Activity> EventList = new List<Activity>();
            List<Activity> CampaignList = new List<Activity>();   //check if it is using
            string offset = string.Empty;
            bool has_more = true;
            while (has_more)
            {
                //maximum load limit is 100 in case of recently modified  engagments 
                string url = "https://api.hubapi.com/email/public/v1/events?hapikey=" + API_Key + "&limit=1000&startTimestamp=" + starttime + "&endTimestamp=" + endtime + "&excludeFilteredEvents=true" + "&offset=" + offset;
                HttpWebRequest webrequest = (HttpWebRequest)WebRequest.Create(url);
                webrequest.Method = "GET";
                webrequest.ContentType = "application/json";
                HttpWebResponse webresponse = null;
                try
                {
                    webresponse = (HttpWebResponse)webrequest.GetResponse();
                }

                catch (Exception e)
                {
                    Console.WriteLine("Get All Engagment exception details:" + e.StackTrace);
                }
                Encoding enc = System.Text.Encoding.GetEncoding("utf-8");
                StreamReader responseStream = new StreamReader(webresponse.GetResponseStream(), enc);
                string result = string.Empty;
                result = responseStream.ReadToEnd();
                webresponse.Close();
                MarketingEmail obj;
                using (var ms = new MemoryStream(Encoding.Unicode.GetBytes(result)))
                {
                    // Deserialization from JSON
                    DataContractJsonSerializer deserializer = new DataContractJsonSerializer(typeof(MarketingEmail));
                    obj = (MarketingEmail)deserializer.ReadObject(ms);

                }

                Console.WriteLine("Total Fetched Activities from HS: {0} ", obj.events.Count);
                for (int i = 0; i < obj.events.Count; i++)
                {
                    //if (obj.events[i].recipient == "matthew.herpich@websterbankarena.com"  && obj.events[i].type!= "PROCESSED")
                   //if (obj.events[i].recipient == "jharris@hcr-manorcare.com")
                    //{
                    if (obj.events[i].type == "DEFERRED" || obj.events[i].type == "PROCESSED")
                    {
                        continue;
                    }
                    else
                    {
                        //if(obj.events[i].type=="CLICK")
                        //{
                        //    Console.WriteLine("new Click Found");
                        //}
                        string subject = obj.events[i].subject;
                        //Console.WriteLine("subnew new:" + subject);
                        string log = "Campaign Id: " + obj.events[i].emailCampaignId + " subject: " + obj.events[i].subject + " Receipient: " + obj.events[i].recipient + " Event Type: " + obj.events[i].type;
                        CampaignLog(log);

                            //getting subscription status code starts here
                            Event e = obj.events[i];
                        string status = "";
                        if (obj.events[i].portalSubscriptionStatus != null)
                        {
                            status = obj.events[i].portalSubscriptionStatus;
                        }
                        else if (obj.events[i].subscriptions != null)
                        {
                            if(obj.events[i].subscriptions.Count > 0 )
                            {
                                status = obj.events[i].subscriptions[0].status;
                            }
                        }
                        //getting subscription status  code ends here
                        int index = EventList.FindIndex(a => (a.emailCampaignId == obj.events[i].emailCampaignId && a.recipient == obj.events[i].recipient));
                        if (index == -1)
                        {
                            activity = new Activity();
                            activity.id = obj.events[i].id;
                            activity.createdon = obj.events[i].created;
                            activity.emailCampaignId = obj.events[i].emailCampaignId;
                            activity.emailCampaignGroupId = obj.events[i].emailCampaignGroupId;
                            activity.recipient = obj.events[i].recipient;
                            activity.subject = subject;
                            //Console.WriteLine("subject" + activity.subject);
                            activity.processing = new List<Events>();
                            activity.tracking = new List<Events>();
                            // Guid contact_id = getGuidDynamicsContactByEmail(activity.recipient);
                            // if (contact_id != Guid.Empty)
                            //{
                            if (activity.subject == null || activity.subject == string.Empty)
                            {

                                activity.isExistInDynamics = true;
                                activity.tracking.Add(new Events() { eventId = obj.events[i].id, type = obj.events[i].type, category = obj.events[i].category, createdon = obj.events[i].created, dropReason = obj.events[i].dropReason, dropMessage = obj.events[i].dropMessage, portalSubscriptionStatus = status, Subscriptions=obj.events[i].subscriptions });


                            }
                            else
                            {
                                //checking if this campaign already exist then skip it otherwise add it.
                                //Activity res = getGuidMarketingEmailEvevtByEmailCampaignId(obj.events[i].emailCampaignId, obj.events[i].recipient);
                                //if (res.activityid == Guid.Empty)
                                //{
                                    //activity.isExistInDynamics = true;
                                    activity.processing.Add(new Events() { eventId = obj.events[i].id, type = obj.events[i].type, category = obj.events[i].category, createdon = obj.events[i].created, dropReason = obj.events[i].dropReason, dropMessage = obj.events[i].dropMessage, portalSubscriptionStatus = status, Subscriptions = obj.events[i].subscriptions });
                                //}
                                //else
                                //{
                                //    continue;
                                //}


                            }
                            //}

                            //else
                            //{
                            //    continue;
                            //}
                            EventList.Add(activity);
                        }
                        else
                        {

                            //activity.id = obj.events[i].id;

                            //activity.emailCampaignId = obj.events[i].emailCampaignId;
                            //activity.emailCampaignGroupId = obj.events[i].emailCampaignGroupId;
                            //activity.recipient = obj.events[i].recipient;
                            //activity.subject = obj.events[i].subject;
                            // Guid contact_id = getGuidDynamicsContactByEmail(activity.recipient);
                            // if (contact_id != Guid.Empty)
                            //{
                            if (obj.events[i].subject == null || obj.events[i].subject == string.Empty)
                            {
                                // activity.isExistInDynamics = false;
                                EventList[index].tracking.Add(new Events() { eventId = obj.events[i].id, type = obj.events[i].type, category = obj.events[i].category, createdon = obj.events[i].created, dropReason = obj.events[i].dropReason, dropMessage = obj.events[i].dropMessage, portalSubscriptionStatus = status , Subscriptions = obj.events[i].subscriptions });
                                //  EventList[index].tracking.Add(activity);
                            }
                            else
                            {
                                EventList[index].subject = obj.events[i].subject;
                                //activity.isExistInDynamics = false;
                                EventList[index].processing.Add(new Events() { eventId = obj.events[i].id, type = obj.events[i].type, category = obj.events[i].category, createdon = obj.events[i].created, dropReason = obj.events[i].dropReason, dropMessage = obj.events[i].dropMessage, portalSubscriptionStatus = status, Subscriptions = obj.events[i].subscriptions });

                            }
                            //}
                            //else
                            //{
                            //    continue;
                            //}

                        }



                    }
                    //}
                }
                has_more = obj.hasMore;
                offset = obj.offset;
                Console.WriteLine("Parameter has-more={0}, offset={1} ", obj.hasMore, obj.offset);
                Console.WriteLine("Preparing  Activities list");
                Console.WriteLine("Please wait it will take some time......");

            }
            int total_before_filteration = EventList.Count;
            int total_unique_contacts = 0;
            int filtered_at_contact_level = 0;
            int total_before_contact_filter = 0;
            int total_after_contact_filter = 0;
            int filtered_at_activity_level = 0;
            int total_update_activities = 0;
            int final_total_activities = 0;
            writeEmailsBeforeFilterationLog(EventList);
            if (EventList.Count == 0)
            {
                return new List<Activity>();
            }
            // Console.WriteLine("Total prepared Activities: " + list.Count);
            List<String> emailsList = EventList.Select(x => x.recipient).Distinct().ToList<String>();
            List<ContactInfo> foundEmails = getGuidDynamicsMultipleContactByEmail(emailsList);
            total_unique_contacts = foundEmails.Count;
            // List<String> foundEmails = new List<string>();
            // List<ContactInfo> foundEmails1 = new List<ContactInfo>();


            //getting counts for new campaign
            var subject_count = 0;
            foreach (var item in EventList)
            {
                if (item.subject != null)
                {
                    subject_count++;
                }
            }

            Console.WriteLine("Events with Subject: " + subject_count);
            //campaign counts
            //  Guid id = Guid.Empty;
            //foreach (var email in emailsList)
            //{
            //    //Add to other List
            //    Console.WriteLine(email);
            //    id = getGuidDynamicsContactByEmail(email);
            //    if (id != Guid.Empty)
            //    {
            //        // foundEmails.Add(email);
            //        foundEmails1.Add(new ContactInfo() { email = email, contacti_id = id });
            //    }
            //}

            List<int> index_list = new List<int>();

            if (foundEmails.Count > 0)
            {
                for (int i = EventList.Count - 1; i > -1; i--)
                {
                    if (foundEmails.Exists(x => x.email == EventList[i].recipient))
                    {
                        int index = foundEmails.FindIndex(x => x.email == EventList[i].recipient);
                        EventList[i].contact_guid = foundEmails[index].contacti_id;
                    }
                    else
                    {
                        //EventList.Remove(EventList[i]);
                        index_list.Add(i);
                    }
                }
            }
            total_before_contact_filter = EventList.Count;
            for (int i = 0; i < index_list.Count; i++)
            {
                EventList.Remove(EventList[index_list[i]]);
            }
            filtered_at_contact_level = index_list.Count;
            total_after_contact_filter = EventList.Count;
            Console.WriteLine("Event counts after Filteration: " + EventList.Count);
            writeEmailsAfterFilterationLog(EventList);
            IEnumerable<Activity> EventPreviousExists = EventList.Where(e => e.subject == string.Empty || e.subject == null).ToList<Activity>();
            List<Activity> EventsNotExist = new List<Activity>();
            foreach (Activity item in EventPreviousExists)
            {
                Activity res = getGuidMarketingEmailEvevtByEmailCampaignId(item.emailCampaignId, item.recipient);
                if (res.activityid != Guid.Empty)
                {
                    EventsNotExist.Add(item);
                    int index = EventList.FindIndex(x => x.recipient == item.recipient && x.emailCampaignId == item.emailCampaignId);
                    EventList[index].activityid = res.activityid;
                    EventList[index].existingDetails = res.existingDetails;
                    EventList[index].existingSubject = res.existingSubject;
                    total_update_activities++;

                }
                else
                {
                    //put log to write file event data
                    filtered_at_activity_level++;
                    EventList.Remove(item);
                    string filter_campaign = "Recepient: " + item.recipient + " Subject:" + item.subject + " Processing:" + item.processing.ToString() + " Tracking:" + item.tracking.ToString();
                    FilteredCampaignLog(filter_campaign);
                }
            }
            //foreach (Activity item in EventsNotExist)
            //{
            //    EventList.Remove(item);
            //}
            //writeList(EventList);
            //EventList.Sort();
            // List<Activity> preparedList = EventList.OrderByDescending(o => o.createdon).ToList();
            //writeList(EventList);
            final_total_activities = EventList.Count;
            Console.WriteLine("Total Activities before Contact filteration: " + total_before_contact_filter);
            Console.WriteLine("Contact Fileration Removed: " + filtered_at_contact_level);
            Console.WriteLine("Total Activities After Contact filteration: " + total_after_contact_filter);
            Console.WriteLine("Total Activities for update: " + total_update_activities);
            Console.WriteLine("Total Activities Removed at Activities Filteration: " + filtered_at_activity_level);
            Console.WriteLine("Final Events count for creation in dynamics: " + final_total_activities);
            var new_create = total_after_contact_filter - total_update_activities - filtered_at_activity_level;
            Console.WriteLine("New Activities to create are: " + new_create);
            Console.WriteLine("Total Activities for update: " + total_update_activities);
            return EventList;
        }
        public static double getUnixTime(DateTime dt)
        {
            double unixTimestamp = (double)(dt.Subtract(new DateTime(1970, 1, 1))).TotalMilliseconds;
            return unixTimestamp;

        }
        static void Main(string[] args)
        {

            DateTime start_time = System.DateTime.Now;
            Console.WriteLine("Start Time: " + start_time);
            Program prg = new Program();
            Guid owner_id = prg.getGuidDynamicsUserByEmail(ConfigurationManager.AppSettings["CRMUserName"]);
            // var res= prg.getOwnerEmailHubSpotByOwnerId(40229753);
            try
            {
                DateTime startTime = Convert.ToDateTime(ConfigurationManager.AppSettings["StartFromDate"]);
                DateTime endTime = Convert.ToDateTime(ConfigurationManager.AppSettings["EndDate"]);
                DateTime baseDate = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                double startingTime = (long)(startTime.ToUniversalTime() - baseDate).TotalMilliseconds;
                double endingTime = (long)(endTime.ToUniversalTime() - baseDate).TotalMilliseconds;
                double startingTime1 = getUnixTime(startTime.ToUniversalTime());
                double endingTime1 = getUnixTime(endTime.ToUniversalTime());


                Console.WriteLine("converted datetime in milliseconds are: " + startingTime + " and : " + ConfigurationManager.AppSettings["StartFromDate"]);
                Console.WriteLine("Start Time: " + startingTime + " ending time: " + endingTime);
                List<Activity> list = prg.getMarketingEmails(startingTime1, endingTime1);

                Console.WriteLine("Total Events are: " + list.Count);
                ////Console.WriteLine(p.getAllPhoneCallActivities(datetime, organizationService));
                //Console.WriteLine("Main-> Call getAllEngagements {Fetching Activities from HS}");
                //List<PhoneCall> list = prg.getAllEngagements(time);
                //Console.WriteLine("Main->  getAllEngagements done");
                //////create multiple requets
                ExecuteMultipleRequest executeMultipe = new ExecuteMultipleRequest()
                {
                    // Assign settings that define execution behavior: continue on error, return responses.
                    Settings = new ExecuteMultipleSettings()
                    {
                        ContinueOnError = true,
                        ReturnResponses = true
                    },
                    // Create an empty organization request collection.
                    Requests = new OrganizationRequestCollection()
                };

                Console.WriteLine("Start Activities creation process in CRM..........");

                if (list.Count > 0)
                {
                    CustomLogging("============= Logging Start============= ");
                    string processing_events = string.Empty;
                    string tracking_events = string.Empty;
                    //start creating email activities
                    foreach (var activity in list)
                    {
                        if (activity.activityid == Guid.Empty)
                        {
                            processing_events = string.Empty;
                            tracking_events = string.Empty;
                            CustomLogging(activity.recipient);
                            //if (activity.contact_email.Equals("richard.miller@followoz.com") || activity.contact_email.Equals("richard.miller@followoz.com"))
                            //{
                            CreateRequest createReq;
                            Entity track = new Entity("new_marketingemail");
                            track["new_emailcampaignid"] = Convert.ToInt32(activity.emailCampaignId);
                            track["new_emailcampaigngroupid"] = Convert.ToInt32(activity.emailCampaignId);

                            List<Events> procesing = activity.processing.OrderByDescending(o => o.createdon).ToList();
                            List<Events> tracking = activity.tracking.OrderByDescending(o => o.createdon).ToList();

                            foreach (Events item in procesing)
                            {
                                DateTime createdon = prg.convertTimestampToDatetime((long)item.createdon).ToLocalTime();
                                string output = string.Empty;
                                if (item.type == "DROPPED")
                                {

                                    output = item.type + " " + item.dropMessage;

                                }
                                else if (item.type == "STATUSCHANGE")
                                {

                                    output = item.portalSubscriptionStatus;

                                }
                                else
                                {
                                    output = item.type;
                                }
                                processing_events += output + " : " + createdon + "\n";
                            }
                            foreach (Events item in tracking)
                            {
                                DateTime createdon = prg.convertTimestampToDatetime((long)item.createdon).ToLocalTime();
                                string output = string.Empty;
                                if (item.type == "BOUNCE")
                                {

                                    string bounced_category = item.category;
                                    if (prg.SoftBounce.Contains(bounced_category))
                                    {
                                        output = "Soft Bounced";
                                    }
                                    else
                                    {
                                        output = "Hard Bounced";
                                    }

                                }
                                else if (item.type == "DROPPED")
                                {

                                    output = item.type + " " + item.dropMessage;

                                }
                                else if (item.type == "STATUSCHANGE")
                                {

                                    output = item.portalSubscriptionStatus;

                                }
                                else if (item.type == "CLICK")
                                {

                                    output = "Clicked";

                                }
                                else if (item.type == "OPEN")
                                {

                                    output = "Opened";

                                }
                                else
                                {
                                    output = item.type;
                                }
                                tracking_events += output + " : " + createdon + "\n";
                            }
                            track["new_eventid"] = activity.id;
                            string subject = activity.subject + " | ⬜ Sent  ⬜ Opened  ⬜ Clicked ⬜ Hard Bounced ⬜ Unsubscribed ";
                            //setting checkbox for sent 


                            if (activity.processing.FindIndex(e => e.type == "SENT") != -1)
                            {
                                string output = prg.FindandReplace(subject, "⬜ Sent", "✅ Sent");
                                if (output != string.Empty)
                                {
                                    subject = output;
                                    track["new_sent"] = true;
                                }
                            }
                            //setting other events checkboxes
                            if (activity.tracking.FindIndex(e => e.type == "OPEN") != -1)
                            {
                                string output = prg.FindandReplace(subject, "⬜ Opened", "✅ Opened");
                                if (output != string.Empty)
                                {
                                    subject = output;
                                    track["new_opened"] = true;
                                }
                            }



                            if (activity.tracking.FindIndex(e => e.type == "BOUNCE") != -1)
                            {
                                int bounced_index = activity.tracking.FindIndex(e => e.type == "BOUNCE");
                                string bounced_category = activity.tracking[bounced_index].category;
                                if (prg.HardBounce.Contains(bounced_category))
                                {
                                    string output = prg.FindandReplace(subject, "⬜ Hard Bounced", "❌ Hard Bounced");
                                    if (output != string.Empty)
                                    {
                                        subject = output;
                                        track["new_bounced"] = true;
                                    }
                                }

                            }
                            if (activity.tracking.FindIndex(e => e.type == "DROPPED") != -1)
                            {

                                string output = prg.FindandReplace(subject, "⬜ Hard Bounced", "❌ Hard Bounced");
                                if (output != string.Empty)
                                {
                                    subject = output;
                                    track["new_bounced"] = true;
                                }


                            }
                            if (activity.processing.FindIndex(e => e.type == "DROPPED") != -1)
                            {

                                string output = prg.FindandReplace(subject, "⬜ Hard Bounced", "❌ Hard Bounced");
                                if (output != string.Empty)
                                {
                                    subject = output;
                                    track["new_bounced"] = true;
                                }


                            }
                            if (activity.tracking.FindIndex(e => e.type == "STATUSCHANGE") != -1)
                            {

                                int index = activity.tracking.FindIndex(e => e.type == "STATUSCHANGE");
                                if (activity.tracking[index].portalSubscriptionStatus == "UNSUBSCRIBED")
                                {

                                    string output = prg.FindandReplace(subject, "⬜ Unsubscribed", "❌ Unsubscribed");
                                    if (output != string.Empty)
                                    {
                                        subject = output;
                                        track["new_unsubscribed"] = true;
                                    }

                                }

                            }

                            if (activity.processing.FindIndex(e => e.type == "STATUSCHANGE") != -1)
                            {
                                int index = activity.processing.FindIndex(e => e.type == "STATUSCHANGE");
                                if (activity.processing[index].portalSubscriptionStatus == "STATUSCHANGE")
                                {
                                    string output = prg.FindandReplace(subject, "⬜ Unsubscribed", "❌ Unsubscribed");
                                    if (output != string.Empty)
                                    {
                                        subject = output;
                                        track["new_unsubscribed"] = true;
                                    }
                                }


                            }

                            if (activity.tracking.FindIndex(e => e.type == "CLICK") != -1)
                            {

                                string output = prg.FindandReplace(subject, "⬜ Clicked", "✅ Clicked");
                                if (output != string.Empty)
                                {
                                    subject = output;
                                    track["new_clicked"] = true;
                                }
                            }

                            track["subject"] = subject;
                            //track["subject"] = activity.subject+ " | ⬜ Sent  ⬜ Bounce  ⬜ Open  ⬜ Click ";
                            track["new_details"] = tracking_events + processing_events;
                            track["new_recipient"] = activity.recipient;
                            track["overriddencreatedon"] = prg.convertTimestampToDatetime(Convert.ToDouble(activity.createdon));
                            track["modifiedon"] = prg.convertTimestampToDatetime(Convert.ToDouble(activity.createdon));
                            // Entity From = new Entity("activityparty");
                            //From["partyid"] = new EntityReference("systemuser", new Guid("7b2678B79C-0EF0-E911-A99C-000D3A37418B"));
                            Entity To = new Entity("activityparty");
                            To["partyid"] = new EntityReference("contact", activity.contact_guid);
                            EntityReference Regarding = new EntityReference("contact", activity.contact_guid);
                            // EntityReference owner_reference = new EntityReference("systemuser", new Guid("7b2678B79C-0EF0-E911-A99C-000D3A37418B"));
                            track["to"] = new Entity[] { To };
                            track["regardingobjectid"] = Regarding;
                            EntityReference owner = new EntityReference();
                            owner.Id = owner_id;
                            track["ownerid"] = owner;

                            createReq = new CreateRequest()
                            {

                                Target = track
                            };
                            executeMultipe.Requests.Add(createReq);
                        }
                        else
                        {
                            processing_events = string.Empty;
                            tracking_events = string.Empty;
                            CustomLogging(activity.recipient);
                            //if (activity.contact_email.Equals("richard.miller@followoz.com") || activity.contact_email.Equals("richard.miller@followoz.com"))
                            //{
                            UpdateRequest updateReq;
                            Entity track = new Entity("new_marketingemail");
                            track.Id = activity.activityid;
                            List<Events> tracking = activity.tracking.OrderByDescending(o => o.createdon).ToList();
                            foreach (Events item in tracking)
                            {
                                DateTime createdon = prg.convertTimestampToDatetime((long)item.createdon).ToLocalTime();
                                string output = string.Empty;
                                if (item.type == "BOUNCE")
                                {

                                    string bounced_category = item.category;
                                    if (prg.SoftBounce.Contains(bounced_category))
                                    {
                                        output = "Soft Bounced";
                                    }
                                    else
                                    {
                                        output = "Hard Bounced";
                                    }

                                }
                                else if (item.type == "DROPPED")
                                {

                                    output = item.type + " " + item.dropMessage;

                                }
                                else if (item.type == "STATUSCHANGE")
                                {

                                    output = item.portalSubscriptionStatus;
                                }
                                else if (item.type == "CLICK")
                                {

                                    output = "Clicked";

                                }
                                else if (item.type == "OPEN")
                                {

                                    output = "Opened";

                                }

                                else
                                {
                                    output = item.type;
                                }
                                tracking_events += output + " : " + createdon + "\n";
                            }
                            track["new_details"] = tracking_events + activity.existingDetails;
                            //updating checkboxes
                            //setting other events checkboxes
                            string subject = activity.existingSubject;
                            if (activity.processing.FindIndex(e => e.type == "SENT") != -1)
                            {
                                string output = prg.FindandReplace(subject, "⬜ Sent", "✅ Sent");
                                if (output != string.Empty)
                                {
                                    subject = output;
                                    track["new_sent"] = true;
                                }
                            }
                            if (activity.tracking.FindIndex(e => e.type == "OPEN") != -1)
                            {
                                string output = prg.FindandReplace(subject, "⬜ Opened", "✅ Opened");
                                if (output != string.Empty)
                                {
                                    subject = output;
                                    track["new_opened"] = true;
                                }
                            }
                            if (activity.tracking.FindIndex(e => e.type == "BOUNCE") != -1)
                            {

                                int bounced_index = activity.tracking.FindIndex(e => e.type == "BOUNCE");
                                string bounced_category = activity.tracking[bounced_index].category;
                                if (prg.HardBounce.Contains(bounced_category))
                                {
                                    string output = prg.FindandReplace(subject, "⬜ Hard Bounced", "❌ Hard Bounced");
                                    if (output != string.Empty)
                                    {
                                        subject = output;
                                        track["new_bounced"] = true;
                                    }
                                }

                            }
                            if (activity.tracking.FindIndex(e => e.type == "DROPPED") != -1)
                            {

                                string output = prg.FindandReplace(subject, "⬜ Hard Bounced", "❌ Hard Bounced");
                                if (output != string.Empty)
                                {
                                    subject = output;
                                    track["new_bounced"] = true;
                                }

                            }

                            if (activity.tracking.FindIndex(e => e.type == "STATUSCHANGE") != -1)
                            {
                                int index = activity.tracking.FindIndex(e => e.type == "STATUSCHANGE");
                                if (activity.tracking[index].portalSubscriptionStatus == "UNSUBSCRIBED")
                                {
                                    string output = prg.FindandReplace(subject, "⬜ Unsubscribed", "❌ Unsubscribed");
                                    if (output != string.Empty)
                                    {
                                        subject = output;
                                        track["new_unsubscribed"] = true;
                                    }
                                }


                            }
                            if (activity.processing.FindIndex(e => e.type == "STATUSCHANGE") != -1)
                            {
                                int index = activity.processing.FindIndex(e => e.type == "STATUSCHANGE");
                                if (activity.processing[index].portalSubscriptionStatus == "UNSUBSCRIBED")
                                {

                                    string output = prg.FindandReplace(subject, "⬜ Unsubscribed", "❌ Unsubscribed");
                                    if (output != string.Empty)
                                    {
                                        subject = output;
                                        track["new_unsubscribed"] = true;
                                    }
                                }
                            }

                            if (activity.tracking.FindIndex(e => e.type == "CLICK") != -1)
                            {

                                string output = prg.FindandReplace(subject, "⬜ Clicked", "✅ Clicked");
                                if (output != string.Empty)
                                {
                                    subject = output;
                                    track["new_clicked"] = true;
                                }
                            }

                            track["subject"] = subject;

                            updateReq = new UpdateRequest()
                            {

                                Target = track
                            };
                            executeMultipe.Requests.Add(updateReq);
                        }
                    }


                    //code for mulitple batch requests

                    ExecuteMultipleRequest executeBatchRequests = new ExecuteMultipleRequest()
                    {
                        // Assign settings that define execution behavior: continue on error, return responses.
                        Settings = new ExecuteMultipleSettings()
                        {
                            ContinueOnError = true,
                            ReturnResponses = true
                        },
                        // Create an empty organization request collection.
                        Requests = new OrganizationRequestCollection()
                    };
                    int counter = 0;
                    int batchsize = 500;
                    for (int x = 0; x < Math.Ceiling((decimal)executeMultipe.Requests.Count / batchsize); x++)
                    {
                        executeBatchRequests.Requests = new OrganizationRequestCollection();
                        var requests = executeMultipe.Requests.Skip(x * batchsize).Take(batchsize);

                        foreach (var request in requests)
                        {
                            executeBatchRequests.Requests.Add(request);
                        }
                        IOrganizationService service = prg.getConnection();
                        ExecuteMultipleResponse responseWithResults = (ExecuteMultipleResponse)service.Execute(executeBatchRequests);
                        Console.WriteLine("This batch create total activities : " + responseWithResults.Responses.Count);
                        counter = counter + responseWithResults.Responses.Count;
                    }
                    Console.WriteLine("Total Activities Created in Dynamics 365 CRM are: {0}", counter);
                    Console.WriteLine("All activities created/synced in CRM successfully ");
                    DateTime end_time = System.DateTime.Now;
                    Console.WriteLine("End Time: " + end_time);
                    Console.WriteLine("Total Duration: " + end_time.Subtract(start_time));
                    // DateTime duration = end_time - start_time;
                    // Console.WriteLine("Total Duration: " + );

                    CustomLogging("===========End============= ");

                }
                else
                {
                    Console.WriteLine("NO activity found for sync");
                }

                var input = "";
                do
                {
                    Console.WriteLine("Enter 'Yes' or 'Y' to stop process: "); 
                    input = Console.ReadLine();
                    if(input.ToUpper()=="YES" || input.ToUpper()=="Y")
                    {
                        break;
                    }
                } while (true);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught - " + ex.StackTrace);
            }
            Console.ReadKey();
        }



        enum ActivityType
        {
            CALL,
            EMAIL,
            MEETING,
            INCOMING_EMAIL
        }
        string[] SoftBounce = new string[]
        {
            "GREYLISTING",
            "MAILBOX_MISCONFIGURATION",
            "ISP_MISCONFIGURATION",
            "IP_REPUTATION",
            "DOMAIN_REPUTATION",
            "DMARC",
            "DNS_FAILURE",
            "SENDING_DOMAIN_MISCONFIGURATION",
            "TEMPORARY_PROBLEM",
            "TIMEOUT",
            "THROTTLED",
            "UNCATEGORIZED",
            "FILTERED",
            "MAILBOX_FULL"

        };
        string[] HardBounce = new string[]
        {

            "CONTENT",
            "SPAM",
            "UNKNOWN_USER",
            "POLICY",
            "HUBSPOT_GLOBAL_BOUNCE"
        };
    }
}
