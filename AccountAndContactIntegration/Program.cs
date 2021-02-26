using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Serialization.Json;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;
using ContactIntegration;
namespace AccountAndContactIntegration
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
        public string getOwnerEmailHubSpotByOwnerId(double id)
        {
            if(id==-1)
            {
                return string.Empty;
            }
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
        public Guid getCompanyIdDynamicsbywebsite(string website)
        {
            Guid account_id = Guid.Empty;
            //Activity activity = new Activity();
            ConditionExpression condition = new ConditionExpression();
            condition.AttributeName = "websiteurl";
            condition.Operator = ConditionOperator.Like;
            condition.Values.Add('%'+website+'%');


            FilterExpression filter = new FilterExpression();
            filter.Conditions.Add(condition);

            QueryExpression query = new QueryExpression("account");
            query.ColumnSet.AddColumns("accountid");
            query.Criteria.AddFilter(filter);
            try
            {
                IOrganizationService service = getConnection();
                EntityCollection records = service.RetrieveMultiple(query);

                if (records.Entities.Count > 0)
                {
                    Entity record = records.Entities.First();
                    account_id = record.GetAttributeValue<Guid>("accountid");
                    //Console.WriteLine("value at index " + userId);

                }

            }
            catch (Exception e)
            {
                string log = "Exception details in getCompanyIdDynamicsbywebsite for website: " + website;
                ErrorLogging(e, log);
                Console.WriteLine("Exception details in getCompanyIdDynamicsbywebsite for website: {2} Message: {0} StackTrack {1}" + e.Message, e.StackTrace, website);

            }
            return account_id;

        }
        public void creatAccountsinDynamics(List<Account> accounts)
        {

            DateTime start_time = System.DateTime.Now;
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

            Console.WriteLine("Started Account creation process in CRM..........");
            int create_count = 0;
            int update_count = 0;
            if (accounts.Count > 0)
            {
                 CustomLogging("============= Starting Account Creation Process============= ");
                //start creating email activities
                foreach (var record in accounts)
                {
                    //  CustomLogging(activity.contact_email);
                    //if (record.is_exist == true)
                    //{
                    //    UpdateRequest updateReq;
                    //    //if (activity.type == ActivityType.CALL.ToString())
                    //    //{
                    //    Entity account = new Entity("account");
                    //    account["createdon"] = record.createdate;
                    //    account["accountid"] = record.dynamics_account_guid;
                    //    account["name"] = record.name;
                    //    //  account["address1_stateorprovince"] = record.state;
                    //    account["new_migratedfrom"] = new OptionSetValue(100000000);
                    //    account["telephone1"] = record.phone;
                    //    account["websiteurl"] = record.website;
                    //    account["address1_city"] = record.city;
                    //    account["address1_county"] = record.country;
                    //    // account["regardingobjectid"] = activity.regarding;
                    //    EntityReference owner = new EntityReference();
                    //    owner.Id = record.dynamics_owner_guid;
                    //    owner.LogicalName = "systemuser";
                    //    account["ownerid"] = owner;
                    //    updateReq = new UpdateRequest()
                    //    {

                    //        Target = account
                    //    };
                    //    executeMultipe.Requests.Add(updateReq);
                    //}
                    CreateRequest createReq;
                    //if (activity.type == ActivityType.CALL.ToString())
                    //{
                    Entity account = new Entity("account");
                    account["createdon"] = record.createdate;
                    account["name"] = record.name;
                    //  account["address1_stateorprovince"] = record.state;
                    account["new_migratedfrom"] = new OptionSetValue(100000000);
                    account["telephone1"] = record.phone;
                    account["websiteurl"] = "https://"+record.website;
                    account["address1_city"] = record.city;
                    account["address1_county"] = record.country;
                    // account["regardingobjectid"] = activity.regarding;
                    EntityReference owner = new EntityReference();
                    owner.Id = record.dynamics_owner_guid;
                    owner.LogicalName = "systemuser";
                    account["ownerid"] = owner;
                    createReq = new CreateRequest()
                    {

                        Target = account
                    };
                    executeMultipe.Requests.Add(createReq);
                }
                //}

                IOrganizationService service = getConnection();
                ExecuteMultipleResponse responseWithResults = (ExecuteMultipleResponse)service.Execute(executeMultipe);
                foreach(var response in responseWithResults.Responses)
                {
                    if(response!=null)
                    {
                        if(response.Response.ResponseName!=null)
                        {
                            if (response.Response.ResponseName.ToString().ToUpper() == "CREATE")
                            {
                                create_count++;
                            }
                            else if (response.Response.ResponseName.ToString().ToUpper() == "UPDATE")
                            {
                                update_count++;
                            }
                        }
                       
                    }

                }
                Console.WriteLine("Total created Accounts are: " + create_count);
                Console.WriteLine("Total updated Accounts are: " + update_count);
                Console.WriteLine("Total Accounts are Created in Dynamics 365 CRM are: {0}", responseWithResults.Responses.Count);
                Console.WriteLine("All Accounts are created/synced in CRM successfully ");
                DateTime end_time = System.DateTime.Now;
                Console.WriteLine("End Time: " + end_time);
                Console.WriteLine("Total Duration: " + end_time.Subtract(start_time));

                   CustomLogging("===========Endind Account Creation Process============= ");
            }
            else
            {
                Console.WriteLine("No Account found for sync");
            }
        }

        public void creatContactsinDynamics(List<ContactInfo> contacts)
        {

            DateTime start_time = System.DateTime.Now;
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

           
            int create_count = 0;
            int update_count = 0;
            if (contacts.Count > 0)
            {
                Console.WriteLine("Started Account creation process in CRM..........");
                CustomLogging("============= Starting Account Creation Process============= ");
                //start contact creating
                foreach (var record in contacts)
                {
                    
                    CreateRequest createReq;
                    Entity contact = new Entity("contact");
                    EntityReference owner = new EntityReference();
                    owner.Id = record.owner_guid;
                    owner.LogicalName = "systemuser";

                    EntityReference regarding = new EntityReference();
                    regarding.Id = record.Account_guid;
                    owner.LogicalName = "customer";

                    //contact["createdon"] = record.createdate;
                    contact["firstname"] = record.firstname;
                    contact["lastname"] = record.lastname;
                    contact["emailaddress1"] = record.email;
                    //contact["websiteurl"] = "https://" + record.;
                    contact["telephone1"] = record.phone;
                    //contact["new_lifecyclestatus"] = record.city;
                  //  contact["new_lifecyclestage"] = record.country;
                   // contact["new_lifecyclestatus"] = record.city;
                    contact["jobtitle"] = record.country;
                   // account["new_migratedfrom"] = record.;
                    //contact["new_campaignsource"] = record.country;
                    contact["aaj_importname"] = record.import_name;
                    contact["aaj_ext"] = record.ext_;
                    //contact["address1_line1"] = record.add;
                    // contact["address1_stateorprovince"] = activity.regarding;
                    //contact["address1_county"] = record.country;
                    //contact["address1_city"] = record.city;
                    //contact["address1_composite"] = record.country;
                    // account["new_originatingsource"] = activity.regarding;


                    // contact["ownerid"] = owner;
                    //contact["parentcustomerid"] = "customer:" + record.Account_guid;
                    contact["parentcustomerid"] = new EntityReference("account", record.Account_guid);
                    createReq = new CreateRequest()
                    {

                        Target = contact
                    };
                    executeMultipe.Requests.Add(createReq);
                }
                
                IOrganizationService service = getConnection();
                ExecuteMultipleResponse responseWithResults = (ExecuteMultipleResponse)service.Execute(executeMultipe);
                foreach (var response in responseWithResults.Responses)
                {
                    if (response != null)
                    {
                        if (response.Response.ResponseName != null)
                        {
                            if (response.Response.ResponseName.ToString().ToUpper() == "CREATE")
                            {
                                create_count++;
                            }
                            else if (response.Response.ResponseName.ToString().ToUpper() == "UPDATE")
                            {
                                update_count++;
                            }
                        }

                    }

                }
                Console.WriteLine("Total created Accounts are: " + create_count);
                Console.WriteLine("Total updated Accounts are: " + update_count);
                Console.WriteLine("Total Accounts are Created in Dynamics 365 CRM are: {0}", responseWithResults.Responses.Count);
                Console.WriteLine("All Accounts are created/synced in CRM successfully ");
                DateTime end_time = System.DateTime.Now;
                Console.WriteLine("End Time: " + end_time);
                Console.WriteLine("Total Duration: " + end_time.Subtract(start_time));

                CustomLogging("===========Endind Account Creation Process============= ");
            }
            else
            {
                Console.WriteLine("No Account found for sync");
            }
        }
        public List<Account> getAllCompanies(double time)
        {
            List<Account> accounts = new List<Account>();
            List<Account> incomplete_accounts = new List<Account>();
            Dictionary<string, Guid> owners = new Dictionary<string, Guid>();

            //declaring variable needed
            string name = string.Empty;
            string domain = string.Empty;
            string website = string.Empty;
            string type = string.Empty;
            string migrate_from = string.Empty;
            string phone = string.Empty;
            string city = string.Empty;
            string state = string.Empty;
            string country = string.Empty;
            double created_date;
            string description = string.Empty;
            string zip = string.Empty;
            string address = string.Empty;
            string address2 = string.Empty;
            string industry = string.Empty;
            string owner_email = string.Empty;
            Guid owner_guid = Guid.Empty;
            Guid account_id = Guid.Empty;
            double owner_id = -1;
            IOrganizationService service = getConnection();
            int offset = 0;
            bool has_more = true;
            while (has_more)
            {
                //maximum load limit is 100 in case of recently modified  engagments 
                string url = "https://api.hubapi.com/companies/v2/companies/recent/created?hapikey=" + API_Key + "&count=100&offset=" + offset+ "&properties=hubspot_owner_id";
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
                   string log = "Exception caught in getAllCompanies function";
                    ErrorLogging(e, log);
                    Console.WriteLine("Exception caught in getAllCompanies function with details" + e.Message);

                }
                Encoding enc = System.Text.Encoding.GetEncoding("utf-8");
                StreamReader responseStream = new StreamReader(webresponse.GetResponseStream(), enc);
                string result = string.Empty;
                result = responseStream.ReadToEnd();
                webresponse.Close();
                AllCompanies obj;
                using (var ms = new MemoryStream(Encoding.Unicode.GetBytes(result)))
                {
                    // Deserialization from JSON
                    DataContractJsonSerializer deserializer = new DataContractJsonSerializer(typeof(AllCompanies));
                    obj = (AllCompanies)deserializer.ReadObject(ms);
                }
                Console.WriteLine("Total Fetched Companies from HS: {0} ", obj.results.Count);
                has_more = obj.hasMore;
                offset = obj.offset;
                Console.WriteLine("Parameter has-more={0}, offset={1} ", obj.hasMore, obj.offset);
                Console.WriteLine("Preparing  Companies list\nPlease wait it will take some time......");
              //making list of all  
                for (int i = 0; i < obj.results.Count; i++)
                {
                    created_date = obj.results[i].properties.createdate.value;
                    if (created_date < time)
                    {
                        continue;
                    }
                    name = string.Empty;
                    domain = string.Empty;
                    website = string.Empty;
                    type = string.Empty;
                    migrate_from = string.Empty;
                    phone = string.Empty;
                    city = string.Empty;
                    state = string.Empty;
                    country = string.Empty;
                    account_id = Guid.Empty;
                    description = string.Empty;
                    zip = string.Empty;
                    address = string.Empty;
                    address2 = string.Empty;
                    owner_email = string.Empty;
                    owner_guid = Guid.Empty;
                    owner_id = -1;
                    if (obj.results[i].properties.address2 != null)
                    {
                        address2 = obj.results[i].properties.address2.value;
                    }
                    if (obj.results[i].properties.address != null)
                    {
                        address = obj.results[i].properties.address.value;
                    }
                    if (obj.results[i].properties.zip != null)
                    {
                        zip = obj.results[i].properties.zip.value;
                    }
                    if (obj.results[i].properties.description != null)
                    {
                        description = obj.results[i].properties.description.value;
                    }
                    if (obj.results[i].properties.name != null)
                    {
                        name = obj.results[i].properties.name.value;
                    }
                    if (obj.results[i].properties.domain != null)
                    {
                        domain = obj.results[i].properties.domain.value;
                    }
                    if (obj.results[i].properties.website != null)
                    {
                        website = obj.results[i].properties.website.value;
                    }
                    if (obj.results[i].properties.type != null)
                    {
                        type = obj.results[i].properties.type.value;
                    }
                    if (obj.results[i].properties.migrated_from != null)
                    {
                        migrate_from = obj.results[i].properties.migrated_from.value;
                    }
                    if (obj.results[i].properties.phone != null)
                    {
                        phone = obj.results[i].properties.phone.value;
                    }
                    if (obj.results[i].properties.city != null)
                    {
                        city = obj.results[i].properties.city.value;
                    }
                    if (obj.results[i].properties.state != null)
                    {
                        state = obj.results[i].properties.state.value;
                    }
                    if (obj.results[i].properties.country != null)
                    {
                        state = obj.results[i].properties.country.value;
                    }
                    if (obj.results[i].properties.hubspot_owner_id != null)
                    {
                        owner_id = obj.results[i].properties.hubspot_owner_id.value;
                    }
                    owner_email = getOwnerEmailHubSpotByOwnerId(owner_id);
                    if (owner_email != string.Empty)
                    {
                        if(owners.ContainsKey(owner_email))
                        {
                            owner_guid = owners.FirstOrDefault(x => x.Key == owner_email).Value;
                        }
                        else
                        {
                            owner_guid = getGuidDynamicsUserByEmail(owner_email);
                           if(owner_guid!=Guid.Empty)
                            {
                                owners.Add(owner_email, owner_guid);
                            }
                        }     
                    }
                    else
                    {
                        continue;
                    }
                    if (owner_guid == Guid.Empty)
                    {
                        continue;
                    }
                    //only check for duplicate if website/domain has value
                    if(domain!=string.Empty)
                    {
                        account_id = getCompanyIdDynamicsbywebsite(domain);
                        if (account_id != Guid.Empty)
                        {
                            Console.WriteLine("Duplicate Account found with this website: " + domain);
                            continue;
                        }
                    }
                    else
                    {
                        Account account1 = new Account { name = name, description = description, domain = domain, website = website, type = type, migrated_from = migrate_from, phone = phone, zip = zip, industry = industry, address = address, address2 = address2, country = country, city = city, createdate = created_date, state = state, dynamics_owner_guid = owner_guid, owner_email = owner_email, is_exist = false, dynamics_account_guid = account_id };
                        incomplete_accounts.Add(account1);
                        continue;
                    }
                    Account account = new Account { name = name, description = description, domain = domain, website = website, type = type, migrated_from = migrate_from, phone = phone, zip = zip, industry = industry, address = address, address2 = address2, country = country, city = city, createdate = created_date, state = state, dynamics_owner_guid=owner_guid, owner_email=owner_email, is_exist=false, dynamics_account_guid=account_id };
                    accounts.Add(account);



                }
            }

            Console.WriteLine("\n\nAccount with Empty websites Details\n\n\n****************************************");
            //printing all accounts which are with empth website
            foreach (var acc in incomplete_accounts)
            {
                Console.WriteLine("Account Missing website with name: " + acc.name);
            }
            Console.WriteLine("****************************************");

            return accounts;
        }


        public string getCompanyWebsiteById(double id)
        {

            string website = string.Empty;
                //maximum load limit is 100 in case of recently modified  engagments 
                string url = "https://api.hubapi.com/companies/v2/companies/"+id+"?hapikey=" + API_Key ;
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
                    string log = "Exception caught in getCompanyWebsiteById function";
                    ErrorLogging(e, log);
                    Console.WriteLine("Exception caught in getCompanyWebsiteById function with details" + e.Message);

                }
                Encoding enc = System.Text.Encoding.GetEncoding("utf-8");
                StreamReader responseStream = new StreamReader(webresponse.GetResponseStream(), enc);
                string result = string.Empty;
                result = responseStream.ReadToEnd();
                webresponse.Close();
                Result obj;
                using (var ms = new MemoryStream(Encoding.Unicode.GetBytes(result)))
                {
                    // Deserialization from JSON
                    DataContractJsonSerializer deserializer = new DataContractJsonSerializer(typeof(Result));
                    obj = (Result)deserializer.ReadObject(ms);
                }
                Console.WriteLine("Company id found for this webiste{1} : {0} ", obj.companyId,website);
            //making list of all  

            if(obj.properties.website!=null)
            {
                website = obj.properties.website.value; 
            }
            return website;

         
        }
        public List<ContactInfo> getAllContacts(double time)
        {
            List<ContactInfo> contacts = new List<ContactInfo>();
            Dictionary<string, Guid> owners = new Dictionary<string, Guid>();

            //declaring variable needed
            string firstname = string.Empty;
            double lastmodifieddate = -1;
            double created_date = -1;
            string lifecyclestage = string.Empty;
            string email = string.Empty;
            string lastname = string.Empty;
            string originating_source = string.Empty;
            string jobtitle = string.Empty;
            string phone = string.Empty;
            string lead_source=string.Empty;
            string company = string.Empty;
            string country = string.Empty;
            string city = string.Empty;
            string import_name = string.Empty;
            string icp_prefix_title = string.Empty;
            string ext_ = string.Empty;
            string lifecycle_status = string.Empty;
            double company_id = -1;
            Guid owner_guid = Guid.Empty;
            Guid account_id = Guid.Empty;
            double owner_id = -1;
            IOrganizationService service = getConnection();
            string vid_offset = string.Empty;
            string vid_time = string.Empty;
            bool has_more = true;
            while (has_more)
            {
                //maximum load limit is 100 in case of recently modified  engagments 
                string url = "https://api.hubapi.com/contacts/v1/lists/all/contacts/recent?hapikey=" + API_Key + "&count=100&property=lifecycle_status&property=lifecyclestage&property=migrated_from&property=firstname&property=lastname&property=jobtitle&property=email&property=phone&property=company&property=address&property=city&property=state&property=country&property=website&property=originating_source&property=icp_prefix_title&property=import_name&property=ext_&property=lead_source&property=hubspot_owner_id&vidOffset=" + vid_offset+ "&timeOffset=" + vid_time;
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
                    Console.WriteLine("Get All companies exception details:" + e.StackTrace);
                }
                Encoding enc = System.Text.Encoding.GetEncoding("utf-8");
                StreamReader responseStream = new StreamReader(webresponse.GetResponseStream(), enc);
                string result = string.Empty;
                result = responseStream.ReadToEnd();
                webresponse.Close();
                AllContacts obj;
                using (var ms = new MemoryStream(Encoding.Unicode.GetBytes(result)))
                {
                    // Deserialization from JSON
                    DataContractJsonSerializer deserializer = new DataContractJsonSerializer(typeof(AllContacts));
                    obj = (AllContacts)deserializer.ReadObject(ms);

                }




                Console.WriteLine("Total Fetched Companies from HS: {0} ", obj.contacts.Count);

                has_more = obj.HasMore;
                vid_offset = Convert.ToString(obj.VidOffset);
                vid_time = Convert.ToString(obj.TimeOffset);
                Console.WriteLine("Parameter has-more={0}, offset={1} ", obj.HasMore, obj.VidOffset);
                Console.WriteLine("Preparing  Activities list");
                Console.WriteLine("Please wait it will take some time......");

                //maing list of all 
                for (int i = 0; i < obj.contacts.Count; i++)
                {

                    //    //check if skip if activity is updated now
                    created_date = obj.contacts[i].properties.createdate.value;
                    if (created_date < time)
                    {
                        continue;
                    }
                    //end update skip check

                    firstname = string.Empty;
                    lastmodifieddate = -1;
                    lifecyclestage = string.Empty;
                    email = string.Empty;
                    lastname = string.Empty;
                    originating_source = string.Empty;
                    jobtitle = string.Empty;
                    phone = string.Empty;
                    lead_source = string.Empty;
                    company_id = -1;
                    account_id = Guid.Empty;
                    company = string.Empty;
                    country = string.Empty;
                    city = string.Empty;
                    import_name = string.Empty;
                    icp_prefix_title = string.Empty;
                    ext_ = string.Empty;
                    lifecycle_status = string.Empty;
                    owner_guid = Guid.Empty;
                    owner_id = -1;

                    if (obj.contacts[i].properties.firstname != null)
                    {
                        firstname = obj.contacts[i].properties.firstname.value;
                    }
                    if (obj.contacts[i].properties.lifecyclestage != null)
                    {
                        lifecyclestage = obj.contacts[i].properties.lifecyclestage.value;
                    }
                    if (obj.contacts[i].properties.email != null)
                    {
                        email = obj.contacts[i].properties.email.value;
                    }
                    if (obj.contacts[i].properties.lastname != null)
                    {
                        lastname = obj.contacts[i].properties.lastname.value;
                    }
                    if (obj.contacts[i].properties.originating_source != null)
                    {
                        originating_source = obj.contacts[i].properties.originating_source.value;
                    }
                    if (obj.contacts[i].properties.jobtitle != null)
                    {
                        jobtitle = obj.contacts[i].properties.jobtitle.value;
                    }
                    if (obj.contacts[i].properties.phone != null)
                    {
                        phone = obj.contacts[i].properties.phone.value;
                    }
                    if (obj.contacts[i].properties.lead_source != null)
                    {
                        lead_source = obj.contacts[i].properties.lead_source.value;
                    }
                    if (obj.contacts[i].properties.company != null)
                    {
                        company = obj.contacts[i].properties.company.value;
                    }
                    if (obj.contacts[i].properties.country != null)
                    {
                        country = obj.contacts[i].properties.country.value;
                    }
                    if (obj.contacts[i].properties.city != null)
                    {
                        city = obj.contacts[i].properties.city.value;
                    }
                    if (obj.contacts[i].properties.import_name != null)
                    {
                        import_name = obj.contacts[i].properties.import_name.value;
                    }

                    if (obj.contacts[i].properties.icp_prefix_title != null)
                    {
                        icp_prefix_title = obj.contacts[i].properties.icp_prefix_title.value;
                    }

                    if (obj.contacts[i].properties.ext_ != null)
                    {
                        ext_ = obj.contacts[i].properties.ext_.value;
                    }

                    if (obj.contacts[i].properties.lifecycle_status!= null)
                    {
                        lifecycle_status = obj.contacts[i].properties.lifecycle_status.value;
                    }
                    if (obj.contacts[i].properties.hubspot_owner_id != null)
                    {
                        owner_id = obj.contacts[i].properties.hubspot_owner_id.value;
                    }
                    string owner_email = getOwnerEmailHubSpotByOwnerId(owner_id);

                    if (owner_email != string.Empty)
                    {
                        if (owners.ContainsKey(owner_email))
                        {
                            owner_guid = owners.FirstOrDefault(x => x.Key == owner_email).Value;

                        }
                        else
                        {
                            owner_guid = getGuidDynamicsUserByEmail(owner_email);
                            if (owner_guid != Guid.Empty)
                            {
                                owners.Add(owner_email, owner_guid);
                            }
                        }

                    }
                    else
                    {
                        continue;
                    }
                   if (owner_guid == Guid.Empty)
                    {
                     continue;
                    }

                    ////only check for duplicate if website/domain has value
                    if (email == string.Empty)
                    {
                        //skiping contact if email is empty
                        continue;
                    }
                   


                    // if (company != string.Empty)
                    //{
                    int contact_id = obj.contacts[i].vid;
                        company_id = getContactAssociation(contact_id);
                   // }
                   if(company_id!=-1)
                    {

                        string company_website = getCompanyWebsiteById(company_id);
                        if(company_website!=string.Empty)
                        {
                            account_id = getCompanyIdDynamicsbywebsite(company_website);
                            if(account_id!=Guid.Empty)
                            {
                                Console.WriteLine("account id is:" + account_id);
                                
                            }
                            else
                            {
                                //related account not found in Dynamics 365 crm
                                Console.WriteLine("Related Account not found in Dynamics");
                                continue;
                            }
                        }
                        else
                        {
                            //skipping if company's website is empty in hubspot
                            continue;
                        }
                       
                    }
                   else
                    {
                        //skipping contact without company associated
                        continue;
                    }
                 

                    contacts.Add(new ContactInfo() { firstname = firstname, lastname = lastname, city = city, company = company, country = country, email = email, ext_ = ext_, jobtitle = jobtitle, phone = phone, icp_prefix_title = icp_prefix_title, import_name = import_name, lead_source = lead_source, lifecyclestage = lifecyclestage, lifecycle_status = lifecycle_status, originating_source = originating_source, createdate = created_date, lastmodifieddate = lastmodifieddate, owner_guid=owner_guid, company_id=Convert.ToDouble(company_id), Account_guid=account_id });


                }
            }
            List<String> emailsList = contacts.Select(x => x.email).Distinct().ToList<String>();
            List<String> list = emailsList.Where(x => x.ToString() != string.Empty).ToList<String>();
            Dictionary<string, Guid> foundEmails = getGuidDynamicsMultipleContactByEmail(list);
            List<int> index_list = new List<int>();
            foreach (ContactInfo contact in contacts)
            {
                if(foundEmails.ContainsKey(contact.email))
                {
                    int index = contacts.FindIndex(x => x.email == contact.email);
                    index_list.Add(index);

                }

            }

            for (int i = 0; i < index_list.Count; i++)
            {
                contacts.Remove(contacts[index_list[i]]);
            }

            return contacts;

        }
        public Dictionary<string,Guid> getGuidDynamicsMultipleContactByEmail(List<string> list)
        {
            Dictionary<string, Guid> contacts = new Dictionary<string, Guid>();
            //int iterations = Convert.ToInt32(list.Count / 2000) + 1;
            List<AccountInfo> output = new List<AccountInfo>();
            //starting from here the batch code
            int counter = 0;
            int batchsize = 1500;
            for (int x = 0; x < Math.Ceiling((decimal)list.Count / batchsize); x++)
            {
                ConditionExpression condition1 = new ConditionExpression();
                condition1.AttributeName = "emailaddress1";
                condition1.Operator = ConditionOperator.In;
                var websites = list.Skip(x * batchsize).Take(batchsize);
                //making request for contacts
                foreach (var website in websites)
                {
                    condition1.Values.Add(website.ToString());
                }
                FilterExpression filter1 = new FilterExpression();
                filter1.Conditions.Add(condition1);
                QueryExpression query = new QueryExpression("contact");
                query.ColumnSet.AddColumns("contactid", "emailaddress1");
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
                            string email = record.GetAttributeValue<string>("emailaddress1");
                            if (contacts.ContainsKey(email))
                            {
                                continue;
                            }
                            else
                            {
                                contacts.Add(record.GetAttributeValue<string>("emailaddress1"), record.GetAttributeValue<Guid>("contactid"));
                            }
                        }
                      
                    }

                }
                catch (Exception e)
                {
                    string log = "Exception details in getGuidDynamicsMultipleContactByEmail  : with complet list ";
                    ErrorLogging(e, log);
                    Console.WriteLine("Exception details in getGuidDynamicsContactByEmail for Email : {2} Message: {0} StackTrack {1}" + e.Message, e.StackTrace, "with this list");

                }

            }


            return contacts;
            //return new Dictionary<string, Guid>();

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

        public double getContactAssociation(double id)
        {
            double company_id = -1;

            string url = "https://api.hubapi.com/crm-associations/v1/associations/" + id + "/HUBSPOT_DEFINED/1?hapikey=" + API_Key;
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
                Console.WriteLine("getContactAssociation exception details:" + e.StackTrace);
            }
            Encoding enc = System.Text.Encoding.GetEncoding("utf-8");
            StreamReader responseStream = new StreamReader(webresponse.GetResponseStream(), enc);
            string result = string.Empty;
            result = responseStream.ReadToEnd();
            webresponse.Close();
            ContactAssociation obj;
            using (var ms = new MemoryStream(Encoding.Unicode.GetBytes(result)))
            {
                // Deserialization from JSON
                DataContractJsonSerializer deserializer = new DataContractJsonSerializer(typeof(ContactAssociation));
                obj = (ContactAssociation)deserializer.ReadObject(ms);

            }

            if (obj.results.Count > 0)
            {
                company_id = obj.results[0];
            }

            return company_id;

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

            Program p = new Program();
            DateTime start_time = System.DateTime.Now;
            Console.WriteLine("Start Time: " + start_time);
            Program prg = new Program();
            //  prg.convertTimestampToDatetime(1592318113123);
            // var res= prg.getOwnerEmailHubSpotByOwnerId(40229753);
            try
            {
                DateTime datetime = Convert.ToDateTime(ConfigurationManager.AppSettings["StartFromDate"]);
                DateTime baseDate = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                var time = (long)(datetime.ToUniversalTime() - baseDate).TotalMilliseconds;
                //Console.WriteLine("converted datetime in milliseconds are: " + time + "and : " + ConfigurationManager.AppSettings["StartFromDate"]);
                //Console.WriteLine("Main-> Call getAllEngagements {Fetching Activities from HS}");
             //  List<Account> accounts = p.getAllCompanies(time);
             //  p.creatAccountsinDynamics(accounts);
                List<ContactInfo> contacts= p.getAllContacts(time);
                p.creatContactsinDynamics(contacts);
                DateTime end_time = System.DateTime.Now;
                Console.WriteLine("End Time: " + end_time);
                Console.WriteLine("Total Duration: " + end_time.Subtract(start_time));

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
           // Console.ReadKey();
        }
    }
}
