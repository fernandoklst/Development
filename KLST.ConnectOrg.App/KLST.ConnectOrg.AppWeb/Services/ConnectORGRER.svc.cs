using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using Microsoft.SharePoint.Client.EventReceivers;

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
                    //result.Status = SPRemoteEventServiceStatus.CancelWithError;
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

                    List<string> tos = new List<string>();
                    try
                    {
                        FieldUserValue[] fTo = itemEvent["ParticipantsPicker"] as FieldUserValue[];
                        foreach(FieldUserValue fuv in fTo)
                        {
                            var userTo = clientContext.Web.EnsureUser(fuv.LookupValue);
                            clientContext.Load(userTo);
                            clientContext.ExecuteQuery();
                            tos.Add(userTo.Email);
                        }
                        ItemURL = itemEvent["URL"].ToString();
                    }
                    catch
                    {

                    }

                    EmailProperties prop = new EmailProperties();
                    prop.Body = String.Format(@"<html><body><h1>This is a ConnectORG Notification</h1><br>Please see your event <a href='{0}'>here</a></body></html>",ItemURL);
                    prop.Subject = "ConnectORG Notification";
                    prop.From = "fernando@klstinc.com";
                    prop.To = tos.AsEnumerable<string>();
                 
                    prop.To = tos.AsEnumerable<string>();
                    Utility.SendEmail(clientContext, prop);
                    clientContext.ExecuteQuery();

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
    }
}
