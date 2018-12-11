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
        /// Email address of secondary recipient
        /// </summary>
        [Category("Options")]
        [DisplayName("1.Recipient Email")]
        [Description("Email address of secondary recipient")]
        public InArgument<String> RecipientEmail { get; set; }

        /// <summary>
        /// The secondary recipient of the mail as Cc
        /// </summary>
        [Category("Options")]
        [DisplayName("2.Cc")]
        [Description("The secondary recipient of the mail as Cc")]
        public InArgument<String> Cc { get; set; }

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
            string recipientEmail = RecipientEmail.Get(context);
            string cc = Cc.Get(context);
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

            EmailMessage reply = responseMessage.Save();
            //Check if attachment is available
            //If attachments
            if (attachments != null && attachments.Length > 0)
                foreach (string attachment in attachments)
                {
                    reply.Attachments.AddFileAttachment(attachment);
                }

            reply.Update(ConflictResolutionMode.AutoResolve);


            //Adding recipients to mail
            string[] recipients = recipientEmail.Split(';');
            foreach (string recipient in recipients)
            {
                //reply.ReplyTo.Add(recipient);
                reply.ToRecipients.Add(recipient);
            }
            Console.WriteLine(reply.ReplyTo);
            
            //Adding recipients to mail
            string[] ccEmails = cc.Split(';');
            foreach (string recipient in ccEmails)
            {
                reply.CcRecipients.Add(recipient);
            }

            reply.From = sender;
            reply.SendAndSaveCopy();
        }
    }
}
