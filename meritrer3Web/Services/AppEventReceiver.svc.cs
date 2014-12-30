using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;
using System.ServiceModel;
namespace meritrer3Web.Services
{
    public class AppEventReceiver : IRemoteEventService
    {
        /// <summary>
        /// Handles app events that occur after the app is installed or upgraded, or when app is being uninstalled.
        /// </summary>
        /// <param name="properties">Holds information about the app event.</param>
        /// <returns>Holds information returned from the app event.</returns>
        public SPRemoteEventResult ProcessEvent(SPRemoteEventProperties properties)
        {
            SPRemoteEventResult result = new SPRemoteEventResult();

            switch (properties.EventType)
            {
                case SPRemoteEventType.AppInstalled: HandleAppInstalled(properties);
                    break;

                //  case SPRemoteEventType.app: HandleAppUninstalled(properties); break;

                case SPRemoteEventType.ItemAdded: HandleItemAdded(properties); break;
            }

            return result;
        }

        /// <summary>
        /// This method is a required placeholder, but is not used by app events.
        /// </summary>
        /// <param name="properties">Unused.</param>
        public void ProcessOneWayEvent(SPRemoteEventProperties properties)
        {
            throw new NotImplementedException();
        }
        private void HandleAppInstalled(SPRemoteEventProperties properties)
        {

            using (ClientContext clientContext = TokenHelper.CreateAppEventClientContext(properties, false))
            {
                if (clientContext != null)
                {
                    bool rerExists = false;
                    List myList = clientContext.Web.Lists.GetByTitle("Merittest1");
                    clientContext.Load(myList, p => p.EventReceivers);
                    clientContext.ExecuteQuery();

                    foreach (var rer in myList.EventReceivers)
                    {
                        if (rer.ReceiverName == "ItemAddedEvent")
                        {
                            rerExists = true;
                        }
                    }

                    if (!rerExists)
                    {
                        EventReceiverDefinitionCreationInformation receiver = new EventReceiverDefinitionCreationInformation();
                        receiver.EventType = EventReceiverType.ItemAdded;
                        receiver.ReceiverName = "ItemAddedEvent";
                        OperationContext op = OperationContext.Current;
                        receiver.ReceiverUrl = op.RequestContext.RequestMessage.Headers.To.ToString();
                        receiver.Synchronization = EventReceiverSynchronization.Synchronous;
                        myList.EventReceivers.Add(receiver);
                        clientContext.ExecuteQuery();
                    }

                }
            }
        }


        private void HandleItemAdded(SPRemoteEventProperties properties)
        {
            using (ClientContext clientContext = TokenHelper.CreateRemoteEventReceiverClientContext(properties))
            {
                if (clientContext != null)
                {
                    clientContext.Load(clientContext.Web);
                    clientContext.ExecuteQuery();

                    SPRemoteEventResult result = new SPRemoteEventResult();

                    Uri myurl = new Uri(properties.ItemEventProperties.WebUrl);


                    if (clientContext != null)
                    {
                        if (properties.EventType == SPRemoteEventType.ItemAdded)
                        {


                            if (
                      properties.ItemEventProperties.ListTitle.Equals("Merittest1", StringComparison.OrdinalIgnoreCase))
                            {
                                List merit1 = clientContext.Web.Lists.GetByTitle("Merittest1");
                                ListItem item = merit1.GetItemById(
                                 properties.ItemEventProperties.ListItemId);
                                clientContext.Load(item);
                                clientContext.ExecuteQuery();

                                item["Title"] += "\nUpdated by RER " +
                                   System.DateTime.Now.ToLongTimeString();
                                item.Update();
                                clientContext.ExecuteQuery();
                            }


                        }
                    }
                }
            }
        }
    }
}
