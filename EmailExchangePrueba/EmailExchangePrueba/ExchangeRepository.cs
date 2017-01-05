using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Threading;

public class ExchangeRepository

{
    private string password = "YOUR_PASSWORD";
    private string accentureAccount = "ACCENTURE_ACCOUNT";
    private string selectedMailBox = "MAILBOX_NAME";

    private ExchangeService _service;
    public string ExceptionMessage { get; set; }

    public ExchangeRepository()
    {
        _service = new ExchangeService();

        //Provide the account user name
        _service.Credentials = new WebCredentials(accentureAccount, password);

        try
        {
            //serviceInstance.TraceEnabled = true;
            //serviceInstance.TraceFlags = TraceFlags.All;

            //serviceInstance.AutodiscoverUrl("enrique.ramos.vargas@accenture.com", SslRedirectionCallback);
            _service.Url = new Uri("https://email.o365.accenture.com/EWS/Exchange.asmx");

            /*
            FindItemsResults<Item> findResults = _service.FindItems(WellKnownFolderName.Inbox, new ItemView(10));
            foreach (Item item in findResults.Items)
            {
                Console.WriteLine(item.Subject);
            }*/

            // Conectamos Inbox del buzón ROBOTICS
            var mailbox = new Mailbox(selectedMailBox);
            var folderId = new FolderId(WellKnownFolderName.Inbox, mailbox);

            Console.WriteLine("===> Correos leídos del Inbox de ROBOTICS:");

            // Leer 10 correos del inbox
            /*FindItemsResults<Item> findResults = _service.FindItems(folderId, new ItemView(10));
            foreach (Item item in findResults.Items)
            {
                Console.WriteLine(item.Subject);
            }*/

            // Subscribe to streaming notifications in the Robotics Inbox. 
            StreamingSubscription streamingSubscription = _service.SubscribeToStreamingNotifications(
                new FolderId[] { folderId },
                EventType.NewMail
               );

            // Create a streaming connection to the service object, over which events are returned to the client.
            // Keep the streaming connection open for 30 minutes.
            StreamingSubscriptionConnection connection = new StreamingSubscriptionConnection(_service, 30);
            connection.AddSubscription(streamingSubscription);
            connection.OnNotificationEvent += OnNotificationEvent;
            connection.OnDisconnect += OnDisconnect;
            connection.Open();

            // Ver diferentes carpetas dentro de Robotics: procesadas, no procesadas, etc.
            /*Folder rootfolder = Folder.Bind(_service, folderId);
            foreach (Folder folder in rootfolder.FindFolders(new FolderView(100)))
            {
                Console.WriteLine("\nName: " + folder.DisplayName + "\n  Id: " + folder.Id);
            }*/
        }
        catch (Exception ex)
        {
            _service = null;
            ExceptionMessage = ex.Message;
        }
    }

    /// <summary>
    /// On disconnect reopen connection automatically
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="args"></param>
    private void OnDisconnect(object sender, SubscriptionErrorEventArgs args)
    {
        // Cast the sender as a StreamingSubscriptionConnection object.           
        StreamingSubscriptionConnection connection = (StreamingSubscriptionConnection)sender;

        connection.Open();
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="args"></param>
    private void OnNotificationEvent(object sender, NotificationEventArgs args)
    {
        StreamingSubscription subscription = args.Subscription;

        // Loop through all item-related events. 
        foreach (NotificationEvent notification in args.Events)
        {

            switch (notification.EventType)
            {
                case EventType.NewMail:
                    Console.WriteLine("\n————-Mail created:————-");
                    break;
            }

            // Display the notification identifier. 
            if (notification is ItemEvent)
            {
                // The NotificationEvent for an e-mail message is an ItemEvent. 
                ItemEvent itemEvent = (ItemEvent)notification;
                Console.WriteLine("\nItemId: " + itemEvent.ItemId.UniqueId);

                Item item1 = Item.Bind(_service, itemEvent.ItemId);
                Console.WriteLine("==> Mail with subject:\n" + item1.Subject);
            }
        }

    }

    

}
