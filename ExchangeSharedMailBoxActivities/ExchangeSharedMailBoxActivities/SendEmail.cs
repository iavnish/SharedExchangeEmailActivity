using System;
using System.Activities;
using System.ComponentModel;
using Microsoft.Exchange.WebServices.Data;


namespace ExchangeSharedMailBoxActivities
{
    public class SendEmail : CodeActivity
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
        /// Mail Subject as string for sending mail
        /// </summary>
        [Category("Input")]
        [DisplayName("2.Subject")]
        [Description("Subject of mail")]
        [RequiredArgument]
        public InArgument<String> Subject { get; set; }

        /// <summary>
        /// Mail Body string
        /// </summary>
        [Category("Input")]
        [DisplayName("3.Body")]
        [Description("Body Text")]
        [RequiredArgument]
        public InArgument<String> Body { get; set; }

        /// <summary>
        /// Email address of recipient
        /// </summary>
        [Category("Input")]
        [DisplayName("4.Recipients Emails")]
        [Description("Email addresses of recipients, use ; for multiple emails")]
        [RequiredArgument]
        public InArgument<String> RecipientEmail { get; set; }

        /// <summary>
        /// Email of sender or Shared Mailbox
        /// Access is required for shared mailbox
        /// </summary>
        [Category("Input")]
        [DisplayName("5.Sender")]
        [Description("Email of sender or Shared Mailbox")]
        [RequiredArgument]
        public InArgument<String> Sender { get; set; }

        /// <summary>
        /// Bool value for is HTML body content
        /// </summary>
        [Category("Options")]
        [DisplayName("1.IsBodyHTML")]
        [Description("True if body is HTML else False")]
        [DefaultValue(false)]
        public InArgument<bool> IsBodyHTML { get; set; }

        /// <summary>
        /// Copy email address of recipient
        /// </summary>
        [Category("Options")]
        [DisplayName("2.Cc")]
        [Description("The secondary recipients of the mail, use ; for multiple emails")]
        public InArgument<String> Cc { get; set; }

        /// <summary>
        /// Email of sender or Shared Mailbox
        /// Access is required for shared mailbox
        /// </summary>
        [Category("Options")]
        [DisplayName("3.Attachments")]
        [Description("File paths of a attachment as a string array")]
        //public List<InArgument<String>> Attachments { get; set; }
        public InArgument<string[]> Attachments { get; set; }

        /// <summary>
        /// Bcc recipient of the mail
        /// </summary>
        [Category("Options")]
        [DisplayName("4.Bcc")]
        [Description("Bcc recipient of the mail")]
        public InArgument<String> Bcc { get; set; }

        /// <summary>
        ///   This function for this class
        ///   It will send mails
        ///   Having option to have attachment
        /// </summary>
        /// <param name="context"></param>
        protected override void Execute(CodeActivityContext context)
        {
            // getting the input values ************************
            ExchangeService objExchangeService = ObjExchangeService.Get(context);
            string subject = Subject.Get(context);
            string body = Body.Get(context);
            string sender = Sender.Get(context);
            bool isBodyHTML = IsBodyHTML.Get(context);
            string recipientEmail = RecipientEmail.Get(context);
            string cc = Cc.Get(context);
            string bcc = Bcc.Get(context);
            string[] attachments = Attachments.Get(context);

        //***** Sending mail Logic ******
        EmailMessage email = new EmailMessage(objExchangeService);
            //Check for if body is a HTML content
            if (isBodyHTML)
                email.Body = new MessageBody(BodyType.HTML, body);
            else
                email.Body = body;

            // Adding Subject to mail
            email.Subject = subject;

            //Adding recipients to mail
            string[] recipients = recipientEmail.Split(';');
            foreach (string recipient in recipients)
            {
                email.ToRecipients.Add(recipient);
            }

            //If attachments
            if (attachments != null &&  attachments.Length > 0)
                foreach (string attachment in attachments)
                {
                    email.Attachments.AddFileAttachment(attachment);
                }

            //If CC is available
            if (cc != null && cc.Length > 0)
            {
                //Adding recipients to mail
                string[] recipientsCC = cc.Split(';');
                foreach (string recipient in recipientsCC)
                {
                    email.CcRecipients.Add(recipient);
                }
            }

            //If BCC is available
            if (bcc != null && bcc.Length > 0)
            {
                //Adding recipients to mail
                string[] recipientsBcc = bcc.Split(';');
                foreach (string recipient in recipientsBcc)
                {
                    email.BccRecipients.Add(recipient);
                }
            }

            //Sending mail and saving it into sent folder
            email.From = sender;


            FolderView view = new FolderView(10000);
            view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
            view.PropertySet.Add(FolderSchema.DisplayName);
            view.Traversal = FolderTraversal.Deep;
            Mailbox mailbox = new Mailbox(sender);
            FindFoldersResults findFolderResults = objExchangeService.FindFolders(new FolderId(WellKnownFolderName.MsgFolderRoot, mailbox), view);

            foreach (Folder folder in findFolderResults)
            {
                if(folder.DisplayName == "Sent Items")
                {
                    email.SendAndSaveCopy(folder.Id);
                }
            }

                
        }
    }
}
