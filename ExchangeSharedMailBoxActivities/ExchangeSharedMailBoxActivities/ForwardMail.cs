using Microsoft.Exchange.WebServices.Data;
using System;
using System.Activities;
using System.ComponentModel;


namespace ExchangeSharedMailBoxActivities
{
    public class ForwardMail : CodeActivity
    {
            /// <summary>
            /// Exchange Service as input
            /// </summary>
            [Category("Input")]
            [DisplayName("1.Exchange Service")]
            [Description("Exchange Service as input")]
            [RequiredArgument]
            public InArgument<ExchangeService> ObjExchangeService { get; set; }

            /// <summary>
            /// Email of sender or Shared Mailbox
            /// </summary>
            [Category("Input")]
            [DisplayName("2.Sender")]
            [Description("Email of sender or Shared Mailbox")]
            [RequiredArgument]
            public InArgument<String> Sender { get; set; }

            /// <summary>
            /// Mail object
            /// </summary>
            [Category("Input")]
            [DisplayName("3.Mail")]
            [Description("Mail object")]
            [RequiredArgument]
            public InArgument<Item> Mail { get; set; }

            /// <summary>
            /// Email address of secondary recipient
            /// </summary>
            [Category("Input")]
            [DisplayName("4.Recipient Email")]
            [Description("Email address of secondary recipient")]
            public InArgument<String> RecipientEmail { get; set; }

            /// <summary>
            /// Body Text
            /// </summary>
            [Category("Input")]
            [DisplayName("5.Body")]
            [Description("Body Text")]
            [RequiredArgument]
            public InArgument<String> Body { get; set; }

            /// <summary>
            /// Condition for checking whether body is html
            /// </summary>
            [Category("Input")]
            [DisplayName("6.IsBodyHTML")]
            [Description("True if body is HTML else False")]
            [RequiredArgument]
            public InArgument<bool> IsBodyHTML { get; set; }


            /// <summary>
            /// The secondary recipient of the mail as Cc
            /// </summary>
            [Category("Options")]
            [DisplayName("1.Cc")]
            [Description("The secondary recipient of the mail as Cc")]
            public InArgument<String> Cc { get; set; }

            /// <summary>
            /// Bcc recipient of the mail
            /// </summary>
            [Category("Options")]
            [DisplayName("2.Bcc")]
            [Description("Bcc recipient of the mail")]
            public InArgument<String> Bcc { get; set; }

        /// <summary>
        /// Shared Mail logic for replying to a mail
        /// </summary>
        /// <param name="context"></param>
        protected override void Execute(CodeActivityContext context)
            {
                // ****************** getting the input values ************************
                ExchangeService objExchangeService = ObjExchangeService.Get(context);
                Item mail = Mail.Get(context);
                string body = Body.Get(context);
                bool isBodyHTML = IsBodyHTML.Get(context);
                string recipientEmail = RecipientEmail.Get(context);
                string cc = Cc.Get(context);
                string bcc = Bcc.Get(context);
                string sender = Sender.Get(context);

                //******** Sending mail Logic ******
                EmailMessage email = EmailMessage.Bind(objExchangeService, mail.Id, new PropertySet(ItemSchema.Attachments));
                ResponseMessage forwardMessage = email.CreateForward();

                //Check for if body is a HTML content
                if (isBodyHTML)
                    forwardMessage.BodyPrefix = new MessageBody(BodyType.HTML, body);
                else
                    forwardMessage.BodyPrefix = body;

                //Adding recipients to mail
                string[] emails = recipientEmail.Split(';');
                foreach (string recipient in emails)
                {
                    forwardMessage.ToRecipients.Add(recipient);
                }

                //Adding Cc to mail
                if (cc != null && cc.Length > 0)
                {
                    string[] emailsCc = cc.Split(';');
                    foreach (string recipient in emailsCc)
                    {
                        forwardMessage.CcRecipients.Add(recipient);
                    }
                }

                //Adding Bcc to mail
                if (bcc != null && bcc.Length > 0)
                {
                    string[] emailsBcc = bcc.Split(';');
                    foreach (string recipient in emailsBcc)
                    {
                        forwardMessage.BccRecipients.Add(recipient);
                    }
                }
                

            //Forwarding mail
            {
                    FolderView view = new FolderView(10000);
                    view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
                    view.PropertySet.Add(FolderSchema.DisplayName);
                    view.Traversal = FolderTraversal.Deep;
                    Mailbox mailbox = new Mailbox(sender);
                    FindFoldersResults findFolderResults = objExchangeService.FindFolders(new FolderId(WellKnownFolderName.MsgFolderRoot, mailbox), view);

                    foreach (Folder folder in findFolderResults)
                    {
                        if (folder.DisplayName == "Sent Items")
                        {
                            forwardMessage.SendAndSaveCopy(folder.Id);
                        }
                    }
                }
            }
        }
    }
