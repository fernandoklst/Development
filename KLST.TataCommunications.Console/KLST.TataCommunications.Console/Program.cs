using System;
using System.Configuration;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.Exchange.WebServices.Data;
using System.IO;
using System.Security;

namespace KLST.TataCommunications.Console
{
    class Program
    {
        static void Main(string[] args)
        {
            string ouser = ConfigurationManager.AppSettings["ouser"].ToString();
            string opass = ConfigurationManager.AppSettings["opass"].ToString();
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            System.Console.WriteLine("Starting App ...");
            service.Credentials = new WebCredentials(ouser, opass);

            //service.TraceEnabled = true;
            //service.TraceFlags = TraceFlags.All;

            service.AutodiscoverUrl(ouser, RedirectionUrlValidationCallback);
            System.Console.WriteLine("Getting Emails ...");

            List<EmailMessage> emails = GetAllEmails(service);

            UpdateEventList(emails);

            BatchDeleteEmailItems(service, GetitemIds(emails));

            System.Console.WriteLine("Press any key to continue ...");
            System.Console.ReadKey();
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

        public static Collection<EmailMessage> BatchGetEmailItems(ExchangeService service, Collection<ItemId> itemIds)
        {

            // Create a property set that limits the properties returned by the Bind method to only those that are required.
            PropertySet propSet = new PropertySet(BasePropertySet.IdOnly, EmailMessageSchema.Subject, EmailMessageSchema.ToRecipients);

            // Get the items from the server.
            // This method call results in a GetItem call to EWS.
            ServiceResponseCollection<GetItemResponse> response = service.BindToItems(itemIds, propSet);

            // Instantiate a collection of EmailMessage objects to populate from the values that are returned by the Exchange server.
            Collection<EmailMessage> messageItems = new Collection<EmailMessage>();


            foreach (GetItemResponse getItemResponse in response)
            {
                try
                {
                    Item item = getItemResponse.Item;
                    EmailMessage message = (EmailMessage)item;
                    messageItems.Add(message);
                    // Print out confirmation and the last eight characters of the item ID.
                    System.Console.WriteLine("Found item {0}.", message.Id.ToString().Substring(144));
                }
                catch (Exception ex)
                {
                    System.Console.WriteLine("Exception while getting a message: {0}", ex.Message);
                }
            }

            // Check for success of the BindToItems method call.
            if (response.OverallResult == ServiceResult.Success)
            {
                System.Console.WriteLine("All email messages retrieved successfully.");
                System.Console.WriteLine("\r\n");
            }

            return messageItems;
        }

        public static List<EmailMessage> GetAllEmails(ExchangeService service)
        {
            int offset = 0;
            int pageSize = 50;
            bool more = true;
            ItemView view = new ItemView(pageSize, offset, OffsetBasePoint.Beginning);

            view.PropertySet = PropertySet.IdOnly;
            FindItemsResults<Item> findResults;
            List<EmailMessage> emails = new List<EmailMessage>();

            while (more)
            {
                findResults = service.FindItems(WellKnownFolderName.Inbox, view);
                foreach (var item in findResults.Items)
                {
                    emails.Add((EmailMessage)item);
                }
                more = findResults.MoreAvailable;
                if (more)
                {
                    view.Offset += pageSize;
                }
            }
            PropertySet properties = (BasePropertySet.FirstClassProperties); //A PropertySet with the explicit properties you want goes here
            service.LoadPropertiesForItems(emails, properties);
            return emails;
        }

        public static void BatchDeleteEmailItems(ExchangeService service, Collection<ItemId> itemIds)
        {
            // Delete the batch of email message objects.
            // This method call results in an DeleteItem call to EWS.
            ServiceResponseCollection<ServiceResponse> response = service.DeleteItems(itemIds, DeleteMode.SoftDelete, null, AffectedTaskOccurrence.AllOccurrences);

            // Check for success of the DeleteItems method call.
            // DeleteItems returns success even if it does not find all the item IDs.
            if (response.OverallResult == ServiceResult.Success)
            {
                System.Console.WriteLine("Email messages deleted successfully.\r\n");
            }

            // If the method did not return success, print a message.
            else
            {
                System.Console.WriteLine("Not all email messages deleted successfully.\r\n");
            }
        }

        public static void UpdateEventList(List<EmailMessage> emails)
        {
            string tcluser = ConfigurationManager.AppSettings["tcluser"].ToString();
            string tclpass = ConfigurationManager.AppSettings["tclpass"].ToString();
            string siteURL = ConfigurationManager.AppSettings["siteURL"].ToString();

            using (ClientContext clientContext = new ClientContext(siteURL))
            {
                SecureString passWord = new SecureString();
                foreach (char c in tclpass.ToCharArray()) passWord.AppendChar(c);
                clientContext.Credentials = new SharePointOnlineCredentials(tcluser, passWord);
                Web web = clientContext.Web;
                List extEvents = web.Lists.GetByTitle("Audit Trail");
                clientContext.Load(extEvents);
                foreach (EmailMessage e in emails)
                {
                    string body = String.IsNullOrEmpty(e.Body.Text) ? String.Empty : e.Body.Text;
                    if (body.ToLower().IndexOf("accept") >= 0 || body.ToLower().IndexOf("reject") >= 0)
                    {
                        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                        ListItem oListItem = extEvents.AddItem(itemCreateInfo);
                        if (body.ToLower().IndexOf("accept") >= 0)
                        {
                            oListItem["Title"] = String.Format("{0} accepted the {1} meeting", e.Sender.Name, e.Subject);
                            oListItem["DateCompleted"] = e.DateTimeReceived;
                        }
                        if (body.ToLower().IndexOf("reject") >= 0)
                        {
                            oListItem["Title"] = String.Format("{0} rejected the {1} meeting", e.Sender.Name, e.Subject);
                        }
                        oListItem.Update();
                        clientContext.ExecuteQuery();
                    }
                }
                
            }
        }
        public static Collection<ItemId> GetitemIds(List<EmailMessage> emails)
        {
            Collection<ItemId> ret = new Collection<ItemId>();
            foreach (EmailMessage em in emails)
            {
                ret.Add(em.Id);
            }
            return ret;
        }
    }

}

