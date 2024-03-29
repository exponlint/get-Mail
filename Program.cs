using System;
using System.Globalization;
using System.IO;
using EAGetMail;
using System.Threading;

namespace consoleApp2
{
    class Program
    {
     
      
        static string _generateFileName(int sequence)
        {
            DateTime currentDateTime = DateTime.Now;
            return string.Format("{0}-{1:000}-{2:000}.eml",
                currentDateTime.ToString("yyyyMMddHHmmss", new CultureInfo("en-US")),
                currentDateTime.Millisecond,
                sequence);
        }


       
        static void Main(string[] args)
        {

            while(true)

            {

                deneme();
            }
           
        }


        static void deneme()
        {
            try
            {
                // Create a folder named "inbox" under current directory
                // to save the email retrieved.
                string localInbox = string.Format("{0}\\inbox", Directory.GetCurrentDirectory());
                // If the folder is not existed, create it.
                if (!Directory.Exists(localInbox))
                {
                    Directory.CreateDirectory(localInbox);
                }

                MailServer oServer = new MailServer("imap.outlook.com",
                        "modelislemit1@outlook.com",
                        "dhptmcwmzmjdulik",
                        ServerProtocol.Imap4);


                oServer.SSLConnection = true;
                oServer.Port = 993;


                MailClient oClient = new MailClient("TryIt");
                oClient.Connect(oServer);

                // retrieve unread/new email only
                oClient.GetMailInfosParam.Reset();
                oClient.GetMailInfosParam.GetMailInfosOptions = GetMailInfosOptionType.NewOnly;

                MailInfo[] infos = oClient.GetMailInfos();
                Console.WriteLine("Total {0} unread email(s)\r\n", infos.Length);
                for (int i = 0; i < infos.Length; i++)
                {
                    
                    MailInfo info = infos[i];
                    Console.WriteLine("Index: {0}; Size: {1}; UIDL: {2}",
                        info.Index, info.Size, info.UIDL);
                    Mail oMail = oClient.GetMail(info);
                    Console.WriteLine("From: {0}", oMail.From.ToString());
                    Console.WriteLine("Subject: {0}\r\n", oMail.Subject);
                    
                    string fileName = _generateFileName(i + 1);
                    string fullPath = string.Format("{0}\\{1}", localInbox, fileName);
                    oMail.SaveAs(fullPath, true);


                    
                    //cc ekle
                    int count;
                    MailAddress addr = oMail.ReplyTo;
                    addr = oMail.From;
                    MailAddress[] addrs = oMail.To;
                    addrs = oMail.Cc;
                    count = addrs.Length;
                    for (int j = 0; j < count; j++)
                    {
                        addr = addrs[j];
                        Console.WriteLine("Cc: {0} <{1}>", addr.Name, addr.Address);
                    }


                    Console.WriteLine(oMail.TextBody);
                    Console.WriteLine("////////////////////////////////////////");
                    //uıdlmanagerkeleme
                    Console.WriteLine(fullPath);
                    Console.WriteLine(fileName);






                    if (!info.Read)
                    {
                        oClient.MarkAsRead(info, true);
                    }

                    Console.WriteLine(oMail.From.ToString());
                }


                oClient.Quit();

            }
            catch (Exception ep)
            {
                Console.WriteLine(ep.Message);
            }

            
            Thread.Sleep(5000);


        }
    }
        
    }
