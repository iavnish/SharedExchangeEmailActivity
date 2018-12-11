using System;
using System.Activities;
using System.ComponentModel;
using System.Security;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeSharedMailBoxActivities
{
    public class InitializeExchangeService : CodeActivity
    {
        /// <summary>
        /// Email address of user
        /// </summary>
        [Category("Input")]
        [DisplayName("User Email")]
        [Description("Email address of user")]
        [RequiredArgument]
        public InArgument<String> UserEmail { get; set; }

        /// <summary>
        /// Password of user mailbox
        /// </summary>
        [Category("Input")]
        [DisplayName("Password")]
        [Description("Password of mailbox")]
        [RequiredArgument]
        public InArgument<SecureString> UserPassword { get; set; }

        /// <summary>
        /// Exchange Server URL
        /// </summary>
        [Category("Input")]
        [DisplayName("Exchange Server URL")]
        [Description("Type Exchange Server URL")]
        [RequiredArgument]
        public InArgument<String> Url { get; set; }

        /// <summary>
        /// Exchange Service as output
        /// </summary>
        [Category("Output")]
        [DisplayName("Exchange Service")]
        [Description("Exchange Service as output")]
        public OutArgument<ExchangeService> ObjExchangeService { get; set; }

        /// <summary>
        /// This is main function for this class
        /// It will create Exchange service object as output
        /// </summary>
        /// <param name="context"></param>
        protected override void Execute(CodeActivityContext context)
        {
            // getting the input values ************************
            string userEmailaddress =  UserEmail.Get(context);
            string userEmailPassword = new System.Net.NetworkCredential(string.Empty, UserPassword.Get(context)).Password; 
            string url = Url.Get(context);

            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013);
            service.Credentials = new System.Net.NetworkCredential(userEmailaddress,userEmailPassword, "");
            service.Url = new Uri(url);

            ObjExchangeService.Set(context, service);
        }
    }
}
