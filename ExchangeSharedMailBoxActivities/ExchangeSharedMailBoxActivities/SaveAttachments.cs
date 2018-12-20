using System;
using System.Activities;
using System.ComponentModel;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeSharedMailBoxActivities
{
    public class SaveAttachments : CodeActivity
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
        /// Mail object
        /// </summary>
        [Category("Input")]
        [DisplayName("2.Mail")]
        [Description("Mail object")]
        [RequiredArgument]
        public InArgument<Item> Mail { get; set; }

        /// <summary>
        /// Full path of a directory
        /// </summary>
        [Category("Input")]
        [DisplayName("3.Folder Path")]
        [Description("Full path of a directory")]
        [RequiredArgument]
        public InArgument<String> FolderPath { get; set; }

        /// <summary>
        /// Shared mail logic for downloading attachments
        /// </summary>
        /// <param name="context"></param>
        protected override void Execute(CodeActivityContext context)
        {
            // ****************** getting the input values ************************
            ExchangeService objExchangeService = ObjExchangeService.Get(context);
            Item mail = Mail.Get(context);
            string folderPath = FolderPath.Get(context);


            //******** Downloading the all attachment from mail ******
            EmailMessage email = EmailMessage.Bind(objExchangeService, mail.Id, new PropertySet(ItemSchema.Attachments));
            foreach (Attachment attachment in email.Attachments)
            {
                if (attachment is FileAttachment)
                {
                    FileAttachment fileAttachment = attachment as FileAttachment;
                    // Load the attachment into a folder.
                    fileAttachment.Load(folderPath +"\\"+ fileAttachment.Name);
                }
            }
        }
    }
}
