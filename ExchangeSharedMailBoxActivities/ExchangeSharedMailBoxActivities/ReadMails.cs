using System;
using System.Activities;
using System.ComponentModel;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeSharedMailBoxActivities
{
    /// <summary>
    /// This is class to read mails with and without filter
    /// Having functions to read only unread mails and all mails
    /// Having function to read mail based on subject filter either exact macth or contains
    /// </summary>
    public class ReadMails : CodeActivity
    {
        /// <summary>
        /// Exchange service property which is being used in other function as input
        /// </summary>
        [Category("Input")]
        [DisplayName("1.Exchange Service")]
        [Description("Exchange Service as Input")]
        [RequiredArgument]
        public InArgument<ExchangeService> ObjExchangeService { get; set; }

        /// <summary>
        /// Folder name property contains, which folder it is going to read
        /// Data Type : String
        /// </summary>
        [Category("Input")]
        [DisplayName("3.Folder Name")]
        [RequiredArgument]
        public InArgument<String> FolderName { get; set; }

        /// <summary>
        /// It contains string, which is address of shared mailbox.
        /// </summary>
        [Category("Input")]
        [DisplayName("2.Mailbox")]
        [Description("Shared Mailbox")]
        [RequiredArgument]
        public InArgument<String> MailBoxName { get; set; }

        /// <summary>
        /// Total number of mail, which it is going to read
        /// </summary>
        [Category("Input")]
        [DisplayName("4.Number of Mails")]
        [Description("Number of mails want to read")]
        [RequiredArgument]
        public InArgument<Int32> numberOfMails { get; set; }

        /// <summary>
        /// Bool value for reading only unread mails or all
        /// </summary>
        [Category("Input")]
        [DisplayName("5.Read Only Unread")]
        [Description("Bool values for only unread mails")]
        [RequiredArgument]
        public InArgument<bool> ReadOnlyUnread { get; set; }

        /// <summary>
        /// Subject is string which will be used to filter based on value
        /// </summary>
        [Category("Options")]
        [DisplayName("1.Subject")]
        [DefaultValue(null)]
        [Description("Subject to filter")]
        public InArgument<String> Subject { get; set; }

        /// <summary>
        /// bool value for filtering exact match with subject value or just contains
        /// Defautl : False
        /// </summary>
        [Category("Options")]
        [DisplayName("2.MatchExact")]
        [DefaultValue(false)]
        [Description("Filter on exact match with subject")]
        public InArgument<bool> MatchExactSubject { get; set; }

        /// <summary>
        /// Out parameter which will contain the fecthed mails from mail box
        /// Data Type : Collection<Item>
        /// </summary>
        [Category("Output")]
        [Description("Collection of Item as output")]
        public OutArgument<System.Collections.ObjectModel.Collection<Item>> EmailMessages { get; set; }

        /// <summary>
        /// This function call other function according to the argument supplied.
        /// This is main function for this class
        /// </summary>
        /// <param name="context"></param>
        protected override void Execute(CodeActivityContext context)
        {
            // ***************  getting the input values ************************
            string folderName = FolderName.Get(context);
            string mailBoxName = MailBoxName.Get(context);
            ExchangeService service = ObjExchangeService.Get(context);
            Int32 numOfMails = numberOfMails.Get(context);
            string mailSubject = Subject.Get(context);
            bool IsExact = MatchExactSubject.Get(context);
            bool onlyUnreadMails = ReadOnlyUnread.Get(context);

            if (mailSubject is null)
            {
                EmailMessages.Set(context, ReadMailFromFolder(service, folderName, numOfMails, mailBoxName, onlyUnreadMails));
            }
            else
            {
                EmailMessages.Set(context, ReadMailFromFolderSubjectFilter(service, folderName, numOfMails, mailBoxName, mailSubject, IsExact));
            }

        }

        /// <summary>
        /// Fecthing mails from a folder of mailbox.
        /// It will fetch without any filter
        /// </summary>
        /// <param name="service"></param>
        /// <param name="FolderName"></param>
        /// <param name="numberOfMails"></param>
        /// <param name="mailboxName"></param>
        /// <param name="onlyUnreadMails"></param>
        /// <returns></returns>
        static System.Collections.ObjectModel.Collection<Item> ReadMailFromFolder(ExchangeService service, String FolderName, Int32 numberOfMails, String mailboxName, bool onlyUnreadMails)
        {
            FolderView view = new FolderView(10000);
            view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
            view.PropertySet.Add(FolderSchema.DisplayName);
            view.Traversal = FolderTraversal.Deep;
            Mailbox mailbox = new Mailbox(mailboxName);
            FindFoldersResults findFolderResults = service.FindFolders(new FolderId(WellKnownFolderName.MsgFolderRoot, mailbox), view);

            foreach (Folder folder in findFolderResults)
            {
                //For Folder filter
                if (folder.DisplayName == FolderName)
                {
                    ItemView itemView = new ItemView(numberOfMails);
                    if (onlyUnreadMails)
                    {
                        SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
                        return ((service.FindItems(folder.Id, searchFilter, itemView)).Items);
                    }
                    else
                        return ((service.FindItems(folder.Id, itemView)).Items);
                }
            }
            throw new Exception("Folder is not found");
        }

        /// <summary>
        /// Fecthing mails from a folder of mailbox.
        /// It will fetch based on subject as filter
        /// </summary>
        /// <param name="service"></param>
        /// <param name="FolderName"></param>
        /// <param name="numberOfMails"></param>
        /// <param name="mailboxName"></param>
        /// <param name="mailSubject"></param>
        /// <param name="isExact"></param>
        /// <returns></returns>
        static System.Collections.ObjectModel.Collection<Item> ReadMailFromFolderSubjectFilter(ExchangeService service, String FolderName, Int32 numberOfMails, String mailboxName, String mailSubject, bool isExact)
        {
            FolderView view = new FolderView(10000);
            view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
            view.PropertySet.Add(FolderSchema.DisplayName);
            view.Traversal = FolderTraversal.Deep;
            Mailbox mailbox = new Mailbox(mailboxName);
            FindFoldersResults findFolderResults = service.FindFolders(new FolderId(WellKnownFolderName.MsgFolderRoot, mailbox), view);
            System.Collections.ObjectModel.Collection<Item> Items = new System.Collections.ObjectModel.Collection<Item>();
            foreach (Folder folder in findFolderResults)
            {
                //For Folder filter
                if (folder.DisplayName == FolderName)
                {
                    ItemView itemView = new ItemView(numberOfMails);
                    if (isExact)
                    {
                        int counter = 0;
                        foreach (Item item in (service.FindItems(folder.Id, itemView)).Items)
                        {
                            //Filtering Based on Subject
                            if (item.Subject == (mailSubject) && counter < numberOfMails)
                            {
                                counter++;
                                Items.Add(item);
                            }
                        }
                    }
                    else
                    {
                        int counter = 0;
                        foreach (Item item in (service.FindItems(folder.Id, itemView)).Items)
                        {
                            //Filtering Based on Subject
                            if (item.Subject.Contains(mailSubject) && counter < numberOfMails)
                            {
                                counter++;
                                Items.Add(item);
                            }
                        }
                    }
                    return (Items);
                }
            }
            return (Items);
        }
    }
}
