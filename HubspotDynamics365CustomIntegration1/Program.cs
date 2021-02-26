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
        public IOrganizationService getConnection()
        {
            ClientCredentials clientCredentials1 = new ClientCredentials();
            clientCredentials1.UserName.UserName = ConfigurationManager.AppSettings["CRMUserName"];
            clientCredentials1.UserName.Password = ConfigurationManager.AppSettings["CRMPWS"];

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            IOrganizationService organizationService1 = (IOrganizationService)new OrganizationServiceProxy(new Uri(ConfigurationManager.AppSettings["CRMApiURL"].ToString()),
             null, clientCredentials1, null);

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
        public static bool isInternalContact(string description)
        {
            return description.Contains("@followoz.com");
        }
        public List<PhoneCall> getAllEngagements( double time)
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

                //maing list of all 
                for (int i = 0; i < obj.results.Count; i++)
                {

                    //check if skip if activity is updated now
                    double activity_create_at = obj.results[i].engagement.createdAt;
                    if(activity_create_at < time )
                    {
                        continue;
                    }
                    //end update skip check

                    //here we are getting engagements 
                    var engagementType = obj.results[i].engagement.type;
                    var engagementBody = obj.results[i].engagement.bodyPreview;
                    var engagmenetBodyHTML = obj.results[i].engagement.bodyPreviewHtml;
                    var engagementSubject = obj.results[i].metadata.subject;
                     var meetingTitle = obj.results[i].metadata.title;
                    var disposition = obj.results[i].metadata.disposition;
                    var ownerId = obj.results[i].engagement.ownerId;
                    decimal activity_id = obj.results[i].engagement.id;

                    //checking if activity is created from dynamics plugin so to skip duplication
                    
                    string subject;
                    if(engagementType==ActivityType.CALL.ToString())
                    {
                        subject = engagementBody+disposition;
                    }
                    else if (engagementType == ActivityType.EMAIL.ToString() || engagementType==ActivityType.INCOMING_EMAIL.ToString())
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
                                int test = 0;
                                //Console.WriteLine(ActivityType.CALL.ToString() + " you entered in condition having more than zero contact");
                                int hubspotContactId = item;
                                PhoneCall activity = new PhoneCall ();
                                activity.new_hubspotactivityid = activity_id;
                                activity.type = engagementType;
                                activity.isfromhubspot = true;
                                activity.createdOn = engagement_createdAt;
                                activity.scheduledEnd = convertTimestampToDatetime((long)obj.results[i].engagement.timestamp).ToLocalTime();
                                activity.subject = engagementSubject;
                                activity.description = engagementBody;
                                string email = getConactEmailHubSpotByContactId(hubspotContactId);
                                if (email == string.Empty || email == null)
                                {
                                    continue;
                                }
                                if (isInternalContact(email) == true)
                                {
                                    continue;
                                }
                                Guid contact_id = getGuidDynamicsContactByEmail(email);
                                if(contact_id.Equals(Guid.Empty))
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
                    if (engagementType == ActivityType.EMAIL.ToString() || engagementType==ActivityType.INCOMING_EMAIL.ToString())
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
                                activity.description = engagmenetBodyHTML;
                                string email = getConactEmailHubSpotByContactId(hubspotContactId);
                                if (email == string.Empty || email == null)
                                {
                                    continue;
                                }
                                if (isInternalContact(email)==true)
                                {
                                    continue;
                                }
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
                            var startTime =convertTimestampToDatetime((long)obj.results[i].metadata.startTime).ToLocalTime();
                            var endTime = convertTimestampToDatetime((long)obj.results[i].metadata.endTime).ToLocalTime();
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
                                activity.scheduledEnd = endTime;
                                activity.scheduleStart = startTime;
                                activity.subject = dynamicSubject;
                                activity.description = engagementBody;
                                string email = getConactEmailHubSpotByContactId(hubspotContactId);
                                if (email == string.Empty || email == null)
                                {
                                    continue;
                                }
                                if (isInternalContact(email) == true)
                                {
                                    continue;
                                }
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
            else if (type=="EMAIL") { entity = "email"; }
            else if (type == "MEETING") { entity = "appointment";  }
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
            if(obj.properties.email!=null)
            {
                email = obj.properties.email.value;
            }
            
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
                Console.WriteLine("Exception details for owner id: {2} Message: {0} StackTrack {1}" + e.Message,e.StackTrace,id);
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
            catch(Exception e)
            {
                
                 string log= "Exception details in getGuidDynamicsUserByEmail for Email : "+ email;
                ErrorLogging(e, log);
                Console.WriteLine("Exception details in getGuidDynamicsUserByEmail for Email : {2} Message: {0} StackTrack {1}" + e.Message, e.StackTrace, email);
            }
            return userId;
        }
        public Guid getGuidDynamicsContactByEmail(string email)
        {
            if(email== "ricky.deane@aa.com")
            {
                Console.WriteLine("got it");
            }
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
            catch(Exception e)
            {
                string log = "Exception details in getGuidDynamicsContactByEmail for Email : " + email;
                ErrorLogging(e, log);
                Console.WriteLine("Exception details in getGuidDynamicsContactByEmail for Email : {2} Message: {0} StackTrack {1}" + e.Message, e.StackTrace, email);

            }
            return contactId;

        }
        public DateTime convertTimestampToDatetime(long timestamp)
        {
            DateTime dtDateTime = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddMilliseconds(timestamp).ToLocalTime();
            //TimeSpan timeSpan = TimeSpan.FromMilliseconds(timestamp);
            //DateTime dt = Convert.ToDateTime(timeSpan.ToString());
            return dtDateTime;
        }
        public static void ErrorLogging(Exception ex,string param)
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
                sw.WriteLine("Details: "+param);
                sw.WriteLine("Error Message: " + ex.Message);
                sw.WriteLine("Stack Trace: " + ex.StackTrace);
                sw.WriteLine("Stack Trace: " + ex.InnerException);
                sw.WriteLine("===========End============= " + DateTime.Now);
            }
        }

        public static void CustomLogging(string param)
        {
            string strPath = @"D:\log.txt";
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
        static void Main(string[] args)
        {

            DateTime start_time = System.DateTime.Now;
            Console.WriteLine("Start Time: "+start_time);
            Program prg = new Program();
          //  prg.convertTimestampToDatetime(1592318113123);
          // var res= prg.getOwnerEmailHubSpotByOwnerId(40229753);
                try
                {
                    DateTime datetime = Convert.ToDateTime(ConfigurationManager.AppSettings["StartFromDate"]);
                    DateTime baseDate = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                    var time = (long)(datetime.ToUniversalTime() - baseDate).TotalMilliseconds;
                     Console.WriteLine("converted datetime in milliseconds are: " + time+ "and : "+ ConfigurationManager.AppSettings["StartFromDate"]);
                //Console.WriteLine(p.getAllPhoneCallActivities(datetime, organizationService));
                Console.WriteLine("Main-> Call getAllEngagements {Fetching Activities from HS}");
                List<PhoneCall> list = prg.getAllEngagements(time);
                Console.WriteLine("Main->  getAllEngagements done");
                ////create multiple requets
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

                if (list.Count > 0 )
                {
                    CustomLogging("============= Logging Start============= ");
                    //start creating email activities
                    foreach (var activity in list)
                    {
                        CustomLogging(activity.contact_email);
                        //if (activity.contact_email.Equals("test@nomail.com") || activity.contact_email.Equals("test@nomail.com"))
                        //{
                            CreateRequest createReq;
                            if (activity.type == ActivityType.CALL.ToString())
                            {
                                Entity phonecall = new Entity("phonecall");
                                phonecall["createdon"] = activity.createdOn;
                                phonecall["scheduledend"] = activity.scheduledEnd;
                                phonecall["subject"] = "Activity Created From Hubspot";
                                phonecall["description"] = activity.description;
                                phonecall["new_isfromhubspot"] = activity.isfromhubspot;
                                phonecall["new_hubspotactivityid"] = Convert.ToDecimal(activity.new_hubspotactivityid);
                                phonecall["from"] = new Entity[] { activity.from };
                                phonecall["to"] = new Entity[] { activity.to };
                                phonecall["regardingobjectid"] = activity.regarding;
                            EntityReference owner = new EntityReference();
                            owner.Id = activity.owner_id;
                            phonecall["ownerid"] = owner;
                            createReq = new CreateRequest()
                                {

                                    Target = phonecall
                                };
                            }
                            else if (activity.type == ActivityType.EMAIL.ToString() || activity.type==ActivityType.INCOMING_EMAIL.ToString())
                            {
                                Entity email = new Entity("email");
                                email["new_isfromhubspot"] = true;
                            email["createdon"] = activity.createdOn;
                                email["scheduledend"] = activity.scheduledEnd;
                                email["new_hubspotactivityid"] = Convert.ToDecimal(activity.new_hubspotactivityid);
                                email["subject"] = activity.subject+" (Activity Created From Hubspot)";
                                email["description"] = activity.description;
                                email["to"] = new Entity[] { activity.to };
                                email["from"] = new Entity[] { activity.from };
                                email["regardingobjectid"] = activity.regarding;
                            EntityReference owner = new EntityReference();
                            owner.Id = activity.owner_id;
                            email["ownerid"] = owner;
                            createReq = new CreateRequest()
                                {

                                    Target = email
                                };
                            }
                            else if (activity.type == ActivityType.MEETING.ToString())
                            {
                                Entity appointment = new Entity("appointment");
                                appointment["new_isfromhubspot"] = true;
                            appointment["createdon"] = activity.createdOn;
                            appointment["new_hubspotactivityid"] = Convert.ToDecimal(activity.new_hubspotactivityid);
                                appointment["subject"] = activity.subject+" (Activity Created From Hubspot)";
                                appointment["description"] = activity.description;
                                appointment["scheduledstart"] = activity.scheduleStart;
                                appointment["scheduledend"] = activity.scheduledEnd;
                                //appointment["RequiredAttendees"] = new Entity[] { activity.to };
                                //appointment["from"] = new Entity[] { activity.from };
                                appointment["regardingobjectid"] = activity.regarding;
                            EntityReference owner = new EntityReference();
                            owner.Id = activity.owner_id;
                            appointment["ownerid"] = owner;
                            createReq = new CreateRequest()
                                {

                                    Target = appointment
                                };
                            }
                            else
                            {
                                continue;
                            }

                            executeMultipe.Requests.Add(createReq);

                        }
                    //}




                     
                    IOrganizationService service = prg.getConnection();
                     ExecuteMultipleResponse responseWithResults = (ExecuteMultipleResponse)service.Execute(executeMultipe);
                    Console.WriteLine("Total Activities Created in Dynamics 365 CRM are: {0}", responseWithResults.Responses.Count);
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
                    if (input.ToUpper() == "YES" || input.ToUpper() == "Y")
                    {
                        break;
                    }
                } while (true);
            }
            catch (Exception ex)
                {
                    Console.WriteLine("Exception caught - " + ex.Message);
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
    }
}
