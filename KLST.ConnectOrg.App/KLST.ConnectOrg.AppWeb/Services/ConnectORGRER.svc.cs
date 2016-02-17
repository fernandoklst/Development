using System;
using System.Collections.Generic;
using System.Collections;
using System.Configuration;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using Microsoft.SharePoint.Client.EventReceivers;
using System.IO;
using System.Net;
using System.Net.Mail;

namespace KLST.ConnectOrg.AppWeb.Services
{
    public class ConnectORGRER : IRemoteEventService
    {
        /// <summary>
        /// Handles events that occur before an action occurs, such as when a user adds or deletes a list item.
        /// </summary>
        /// <param name="properties">Holds information about the remote event.</param>
        /// <returns>Holds information returned from the remote event.</returns>
        public SPRemoteEventResult ProcessEvent(SPRemoteEventProperties properties)
        {
            SPRemoteEventResult result = new SPRemoteEventResult();

            switch (properties.EventType)
            {
                case SPRemoteEventType.ItemUpdating:
                    //result.ErrorMessage = "You cannot add this list item";
                    //result.Status = SPRemoteEventServiceStatus.CancelNoError;
                    break;

            }

            using (ClientContext clientContext = TokenHelper.CreateRemoteEventReceiverClientContext(properties))
            {
                if (clientContext != null)
                {
                    clientContext.Load(clientContext.Web);
                    clientContext.ExecuteQuery();
                }
            }

            return result;
        }

        /// <summary>
        /// Handles events that occur after an action occurs, such as after a user adds an item to a list or deletes an item from a list.
        /// </summary>
        /// <param name="properties">Holds information about the remote event.</param>
        public void ProcessOneWayEvent(SPRemoteEventProperties properties)
        {
            using (ClientContext clientContext = TokenHelper.CreateRemoteEventReceiverClientContext(properties))
            {
                if (clientContext != null)
                {
                    //string Attendees = properties.ItemEventProperties.AfterProperties["ParticipantsPicker"].ToString();
                    List lstExternalEvents = clientContext.Web.Lists.GetByTitle(properties.ItemEventProperties.ListTitle);
                    ListItem itemEvent = lstExternalEvents.GetItemById(properties.ItemEventProperties.ListItemId);
                    clientContext.Load(itemEvent);
                    clientContext.ExecuteQuery();

                    clientContext.Load(clientContext.Web);
                    clientContext.ExecuteQuery();

                    string ItemURL = String.Empty;
                    List<Event> oList = new List<Event>();
                    Event oEvent = Event.Parse(itemEvent);
                    oEvent.Organizer = @"tatacalendar@outlook.com";
                    oList.Add(oEvent);
                    string att = oEvent.ToString(oList);
                    
                    

                    List<string> tos = new List<string>();

                    string smtpserver = ConfigurationManager.AppSettings["smtpserver"].ToString();
                    string username = ConfigurationManager.AppSettings["username"].ToString();
                    string password = ConfigurationManager.AppSettings["password"].ToString();

                    SmtpClient mailserver = new SmtpClient(smtpserver, Convert.ToInt32(587));
                    mailserver.Credentials = new NetworkCredential(username, password);
                    MailAddress from = new MailAddress(@"tatacalendar@outlook.com","Shakir John");
                    MailMessage mess = new MailMessage();

                    try
                    {
                        FieldUserValue[] fTo = itemEvent["ParticipantsPicker"] as FieldUserValue[];
                        foreach (FieldUserValue fuv in fTo)
                        {
                            var userTo = clientContext.Web.EnsureUser(fuv.LookupValue);
                            clientContext.Load(userTo);
                            clientContext.ExecuteQuery();
                            tos.Add(userTo.Email);
                            mess.To.Add(new MailAddress(userTo.Email));
                        }
                        ItemURL = String.Empty;//itemEvent["URL"].ToString();
                    }
                    catch(Exception ex)
                    {

                    }

                    MemoryStream ms = new MemoryStream(GetData(att));
                    System.Net.Mime.ContentType ct = new System.Net.Mime.ContentType("text/calendar");
                    System.Net.Mail.Attachment attach = new System.Net.Mail.Attachment(ms, ct);
                    attach.ContentDisposition.FileName = "event.ics";

                    try
                    {
                        string subject = "ConnectORG Notification";
                        string body = String.Format(@"<html><body><h1>This is a ConnectORG Notification</h1><br>Please see your event <a href='{0}'>here</a></body></html>", ItemURL);
                        mess.From = from;
                        mess.Subject = subject;
                        mess.Body = body;
                        mess.IsBodyHtml = true;
                        mess.Attachments.Add(attach);
                        mailserver.Send(mess);
                    }
                    catch(SmtpException ex)
                    {

                    }
                    catch(Exception e)
                    {
                        
                    }

                    try
                    {
                        List AuditTrail = clientContext.Web.Lists.GetByTitle("Audit Trail");
                        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                        ListItem newItem = AuditTrail.AddItem(itemCreateInfo);
                        newItem["Title"] = String.Format("Sent Email to {0} people", tos.Count());
                        newItem["DateCompleted"] = DateTime.Now;
                        newItem.Update();
                        clientContext.ExecuteQuery();
                    }
                    catch
                    {

                    }
                  
                }
            }
        }

        static byte[] GetData(string s)
        {
            //this method just returns some binary data.
            byte[] data = Encoding.ASCII.GetBytes(s);
            return data;
        }
    }
    public class Event
    {
        public int ID { get; set; }
        public string Title { get; set; }
        public string Organizer { get; set; }
        public DateTime Created { get; set; }
        public string Description { get; set; }
        public string Location { get; set; }
        public string Category { get; set; }
        public string UID { get; set; }
        public DateTime EventDate { get; set; }
        public DateTime EndDate { get; set; }
        public int Duration { get; set; }
        public bool Recurrence { get; set; }
        public string RecurrenceData { get; set; }
        public DateTime RecurrenceID { get; set; }
        public int MasterSeriesItemID { get; set; }
        public EventType EventType { get; set; }
        public bool AllDayEvent { get; set; }
        public DateTime LastModified { get; set; }

        public static Event Parse(ListItem sharepointListItem)
        {
            Event e = new Event()
            {
                ID = (int)sharepointListItem["ID"],
                Title = (string)sharepointListItem["Title"],
                Created = DateTime.Parse(sharepointListItem["Created"].ToString()),
                Description = (string)sharepointListItem["Description"],
                Location = (string)sharepointListItem["Location"],
                Category = (string)sharepointListItem["Category"],
                UID = sharepointListItem["UID"] == null ? Guid.NewGuid().ToString() : sharepointListItem["UID"].ToString(),
                EventDate = DateTime.Parse(sharepointListItem["EventDate"].ToString()),
                EndDate = DateTime.Parse(sharepointListItem["EndDate"].ToString()),
                Duration = (int)sharepointListItem["Duration"],
                Recurrence = (bool)sharepointListItem["fRecurrence"],
                RecurrenceData = (string)sharepointListItem["RecurrenceData"],
                RecurrenceID = sharepointListItem["RecurrenceID"] != null ? DateTime.Parse(sharepointListItem["RecurrenceID"].ToString()) : DateTime.MinValue,
                MasterSeriesItemID = sharepointListItem["MasterSeriesItemID"] == null ? -1 : (int)sharepointListItem["MasterSeriesItemID"],
                EventType = (EventType)Enum.Parse(typeof(EventType), sharepointListItem["EventType"].ToString()),
                AllDayEvent = (bool)sharepointListItem["fAllDayEvent"],
                LastModified = DateTime.Parse(sharepointListItem["Last_x0020_Modified"].ToString())
            };

            return e;
        }

        public string ToString(List<Event> Events)
        {
            StringBuilder builder = new StringBuilder();

            builder.AppendLine("BEGIN:VCALENDAR");
            builder.AppendLine("VERSION:2.0");
            builder.AppendLine("METHOD:REQUEST");
            builder.AppendLine("PRODID: -//imc");
            builder.AppendLine("BEGIN:VEVENT");
            builder.AppendLine("SUMMARY:" + CleanText(Title));
            builder.AppendLine("DTSTAMP:" + Created.ToString("yyyyMMddTHHmmssZ"));
            builder.AppendLine("DESCRIPTION:" + CleanText(Description));
            builder.AppendLine("LOCATION:" + CleanText(Location));
            builder.AppendLine("CATEGORIES:" + CleanText(Category));
            builder.AppendLine("ATTENDEE;RSVP=TRUE:mailto:" + CleanText(Organizer));
            builder.AppendLine("ORGANIZER:mailto:" + CleanText(Organizer));
            builder.AppendLine("UID:" + UID);
            builder.AppendLine("STATUS:CONFIRMED");
            builder.AppendLine("LAST-MODIFIED:" + LastModified.ToString("yyyyMMddTHHmmssZ"));

            if (AllDayEvent)
            {
                builder.AppendLine("DTSTART;VALUE=DATE:" + EventDate.ToString("yyyyMMdd"));

                double days = Math.Round(((Double)Duration / (double)(60 * 60 * 24)));
                builder.AppendLine("DTEND;VALUE=DATE:" + EventDate.AddDays(days).ToString("yyyyMMdd"));
            }
            else
            {
                builder.AppendLine("DTSTART:" + EventDate.ToString("yyyyMMddTHHmmssZ"));
                builder.AppendLine("DTEND:" + EventDate.AddSeconds(Duration).ToString("yyyyMMddTHHmmssZ"));
            }

            IEnumerable<Event> deletedEvents = Events.Where(e => e.MasterSeriesItemID == ID && e.EventType == EventType.Deleted);
            foreach (Event deletedEvent in deletedEvents)
            {
                if (AllDayEvent)
                    builder.AppendLine("EXDATE;VALUE=DATE:" + deletedEvent.RecurrenceID.ToString("yyyyMMdd"));
                else
                    builder.AppendLine("EXDATE:" + deletedEvent.RecurrenceID.ToString("yyyyMMddTHHmmssZ"));
            }

            if (RecurrenceID != DateTime.MinValue && EventType == EventType.Exception) //  Event is exception to recurring item
            {
                if (AllDayEvent)
                    builder.AppendLine("RECURRENCE-ID;VALUE=DATE:" + RecurrenceID.ToString("yyyyMMdd"));
                else
                    builder.AppendLine("RECURRENCE-ID:" + RecurrenceID.ToString("yyyyMMddTHHmmssZ"));
            }
            else if (Recurrence && !RecurrenceData.Contains("V3RecurrencePattern"))
            {
                RecurrenceHelper recurrenceHelper = new RecurrenceHelper();
                builder.AppendLine(recurrenceHelper.BuildRecurrence(RecurrenceData, EndDate));
            }

            if (EventType == EventType.Exception)
            {
                List<Event> exceptions = Events.Where(e => e.MasterSeriesItemID == MasterSeriesItemID).OrderBy(e => e.Created).ToList<Event>();
                builder.AppendLine("SEQUENCE:" + (exceptions.IndexOf(this) + 1));
            }
            else
                builder.AppendLine("SEQUENCE:0");

            builder.AppendLine("BEGIN:VALARM");
            builder.AppendLine("ACTION:DISPLAY");
            builder.AppendLine("TRIGGER:-PT10M");
            builder.AppendLine("DESCRIPTION:Reminder");
            builder.AppendLine("END:VALARM");

            builder.AppendLine("END:VEVENT");
            builder.AppendLine("END:VCALENDAR");

            return builder.ToString();
        }

        private string CleanText(string text)
        {
            if (text != null)
                text = text.Replace(@"\", "\\\\")
                            .Replace(";", @"\;")
                            .Replace(",", @"\,")
                            .Replace("\r\n", @"\n");
            return text;
        }
    }

    public enum SPEventType
    {
        Add = 0x00000001, Modify = 0x00000002, Delete = 0x00000004, Discussion= 0x00000FF0, All=-1
    }

    public enum EventType
    {
        Exception = 1, Deleted = 4
    }
}
