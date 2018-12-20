using System;
using System.Activities;
using System.ComponentModel;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeSharedMailBoxActivities
{
    public class ReplyToMail : CodeActivity
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
        /// Body Text
        /// </summary>
        [Category("Input")]
        [DisplayName("4.Body")]
        [Description("Body Text")]
        [RequiredArgument]
        public InArgument<String> Body { get; set; }

        /// <summary>
        /// Condition for checking whether body is html
        /// </summary>
        [Category("Input")]
        [DisplayName("5.IsBodyHTML")]
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
        /// File path of a attachment
        /// </summary>
        [Category("Options")]
        [DisplayName("3.Attachments")]
        [Description("File path of a attachments as string array")]
        public InArgument<string[]> Attachments { get; set; }

        /// <summary>
        /// Filter on exact match with subject
        /// </summary>
        [Category("Options")]
        [DisplayName("4.Reply All")]      
        [Description("Bool value for reply to all")]
        [DefaultValue(false)]
        public InArgument<bool> ReplyAll { get; set; }


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
            //string recipientEmail = RecipientEmail.Get(context);
            string cc = Cc.Get(context);
            string bcc = Bcc.Get(context);
            string sender = Sender.Get(context);
            string[] attachments = Attachments.Get(context);
            bool isReplyAll = ReplyAll.Get(context);

            //******** Sending mail Logic ******
            EmailMessage email = EmailMessage.Bind(objExchangeService, mail.Id, new PropertySet(ItemSchema.Attachments));
            ResponseMessage responseMessage = email.CreateReply(isReplyAll);

            //Check for if body is a HTML content
            if (isBodyHTML)
                responseMessage.BodyPrefix = new MessageBody(BodyType.HTML, body);
            else
                responseMessage.BodyPrefix = body;

            //If CC is available
            if (cc != null && cc.Length > 0)
            {
                //Adding recipients to mail
                string[] recipientsCC = cc.Split(';');
                foreach (string recipient in recipientsCC)
                {
                    responseMessage.CcRecipients.Add(recipient);
                }
            }

            //If BCC is available
            if (bcc != null && bcc.Length > 0)
            {
                //Adding recipients to mail
                string[] recipientsBcc = bcc.Split(';');
                foreach (string recipient in recipientsBcc)
                {
                    responseMessage.BccRecipients.Add(recipient);
                }
            }


            //Check if attachment is available
            //If attachments
            if (attachments != null && attachments.Length > 0)
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

                        //Adding attachments to reply mail
                        EmailMessage reply = responseMessage.Save(folder.Id);
                        foreach (string attachment in attachments)
                        {
                            reply.Attachments.AddFileAttachment(attachment);
                        }
                        reply.Update(ConflictResolutionMode.AlwaysOverwrite);

                        //Sending mail and saving to sent Items
                        reply.SendAndSaveCopy(folder.Id);
                    }
                }

            }

            else
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
                        responseMessage.SendAndSaveCopy(folder.Id);
                    }
                }

            }
        }
    }
}
