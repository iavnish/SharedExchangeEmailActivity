using System;
using System.Activities;
using System.ComponentModel;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeSharedMailBoxActivities
{
    public class MoveMail : CodeActivity
    {
        /// <summary>
        /// Exchange Service as input
        /// </summary>
        [Category("Input")]
        [DisplayName("Exchange Service")]
        [Description("Exchange Service as input")]
        [RequiredArgument]
        public InArgument<ExchangeService> ObjExchangeService { get; set; }

        /// <summary>
        /// Mail object
        /// </summary>
        [Category("Input")]
        [DisplayName("Mail")]
        [Description("Mail object")]
        [RequiredArgument]
        public InArgument<Item> Mail { get; set; }

        /// <summary>
        /// Shared mailbox address
        /// </summary>
        [Category("Input")]
        [DisplayName("Mailbox")]
        [Description("Shared mailbox address")]
        [RequiredArgument]
        public InArgument<String> MailboxName { get; set; }

        /// <summary>
        /// Move to folder name
        /// </summary>
        [Category("Input")]
        [DisplayName("Folder Name")]
        [Description("Move to folder name")]
        [RequiredArgument]
        public InArgument<String> FolderName { get; set; }

        /// <summary>
        /// Logic for moving mails
        /// </summary>
        /// <param name="context"></param>
        protected override void Execute(CodeActivityContext context)
        {
            // ****************** getting the input values ************************
            ExchangeService objExchangeService = ObjExchangeService.Get(context);
            Item mail = Mail.Get(context);
            string mailboxName = MailboxName.Get(context);
            string folderName = FolderName.Get(context);

            //********** Logic to move a mail from a folder to another ************
            FolderView view = new FolderView(10000);
            view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
            view.PropertySet.Add(FolderSchema.DisplayName);
            view.Traversal = FolderTraversal.Deep;
            Mailbox mailbox = new Mailbox(mailboxName);
            FindFoldersResults findFolderResults = objExchangeService.FindFolders(new FolderId(WellKnownFolderName.MsgFolderRoot, mailbox), view);
            bool flag = false;

            foreach (Folder folder in findFolderResults)
            {
                //Searching for supplied folder into mailbox
                if (folder.DisplayName.Equals(folderName))
                {
                    PropertySet propSet = new PropertySet(BasePropertySet.IdOnly, EmailMessageSchema.Subject, EmailMessageSchema.ParentFolderId);
                    // Bind to the existing item by using the ItemId.
                    EmailMessage beforeMessage = EmailMessage.Bind(objExchangeService, mail.Id, propSet);
                    // Move the specified mail to the given folder
                    beforeMessage.Move(folder.Id);
                    flag = true;
                    break;
                }
            }
            if (!flag)
                throw new Exception("Supplied folder is not found");
        }
    }
}
