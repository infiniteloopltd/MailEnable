using System;
using MailEnable.Administration;
using System.Collections.Generic;

namespace MailEnable
{
    class Program
    {
        static void Main()
        {
            AddMailbox("default","john@printfromipad.com","password");
            var mailboxes = ListMailboxes("default");
            foreach (var mailbox in mailboxes)
            {
                Console.WriteLine(mailbox);
            }
            Console.ReadLine();
        }

        private static IEnumerable<string> ListMailboxes(string postOffice)
        {
            var oMailbox = new Mailbox
            {
                Postoffice = postOffice
            };
            if (oMailbox.FindFirstMailbox() != 1)
            {
                throw new Exception("Failed to find mailboxes");
            }
            var mailboxes = new List<string>();
            do
            {
                var sMailboxName = oMailbox.MailboxName;
                mailboxes.Add(sMailboxName);
            }
            while (oMailbox.FindNextMailbox() == 1);
            return mailboxes;
        }

        private static void AddMailbox(string postOffice, string email, string password)
        {
            var user = email.Split('@')[0];
            var domain = email.Split('@')[1];
            var oMailbox = new MailEnable.Administration.Mailbox();
            var oLogin = new Administration.Login();
            var sMailboxName = user;
            var sPassword = password;
            const string sRights = "USER"; // USER – User, ADMIN – Administrator;
            oLogin.Account = postOffice;
            oLogin.LastAttempt = -1;
            oLogin.LastSuccessfulLogin = -1;
            oLogin.LoginAttempts = -1;
            oLogin.Password = "";
            oLogin.Rights = "";
            oLogin.Status = -1;
            oLogin.UserName = sMailboxName + "@" + domain;
            // If the login does not exist we need to create it
            if (oLogin.GetLogin() == 0)
            {
                oLogin.Account = postOffice;
                oLogin.LastAttempt = 0;
                oLogin.LastSuccessfulLogin = 0;
                oLogin.LoginAttempts = 0;
                oLogin.Password = sPassword;
                oLogin.Rights = sRights;
                oLogin.Status = 1; // 0 – Disabled, 1 – Enabled
                oLogin.UserName = sMailboxName + "@" + domain;
                if (oLogin.AddLogin() != 1)
                {
                    // Error adding the Login
                    throw new Exception("Failed to add Login");
                }
            }
            // Now we create the mailbox
            oMailbox.Postoffice = postOffice;
            oMailbox.MailboxName = sMailboxName;
            oMailbox.Size = 0;
            oMailbox.Limit = -1; // -1 – Unlimited OR size value (in KB)
            oMailbox.Status = 1;
            if (oMailbox.AddMailbox() != 1)
                // Failed to add mailbox
                throw new Exception("Failed to add mailbox");
            // Now we need to add the Address Map entries for the Account
            var oAddressMap = new MailEnable.Administration.AddressMap
            {
                Account = postOffice,
                DestinationAddress = "[SF:" + postOffice + "/" + sMailboxName + "]",
                SourceAddress = "[SMTP:" + sMailboxName + "@" + domain + "]",
                Scope = "0" // ?
            };
            if (oAddressMap.AddAddressMap() != 1)
                // Failed to add Address Map for some reason!
                throw new Exception("Failed to add AddressMap");
          
        }

    }
}
