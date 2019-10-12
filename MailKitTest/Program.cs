using ExcelDataReader;
using MailKit.Net.Smtp;
using MimeKit;
using System;
using System.IO;
using System.Text;

namespace MailKitTest
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Start...");
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string filePath = @"C:\Users\joleitne\Desktop\emailAddresses.csv";
            Int64 emailCount = 1;
            string failedEmailAddress = "";
            int retry = 3;

            Console.WriteLine("Attempting to open " + filePath);
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                Console.WriteLine("Opened ");
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx)
                Console.WriteLine("Attempting to read " + filePath);
                using (var reader = ExcelReaderFactory.CreateCsvReader(stream))
                {
                    Console.WriteLine("Read ");
                    // Choose one of either 1 or 2:

                    // 1. Use the reader methods
                    do
                    {
                        while (reader.Read())
                        {
                            Console.WriteLine("Reading Row {0}", emailCount);
                            string emailAddress = reader.GetString(0);
                            try
                            {
                                //Check if the current emailAddress has failed. 
                                if (failedEmailAddress == emailAddress)
                                {
                                    Console.WriteLine("retrying {0}. {1} Attempts left", failedEmailAddress, retry);
                                    retry--;
                                }

                                //have we tried this email address more than three times?
                                if (retry >= 0)
                                {
                                    sendEmail(emailAddress);
                                    Console.WriteLine("Email [ {0} ] Sent", emailCount);
                                    failedEmailAddress = "";
                                    emailCount++;
                                }
                                
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("ERROR: {0} failed to Send",emailAddress);
                                failedEmailAddress = emailAddress;
                                Console.WriteLine(ex.ToString());



                            }

                        }
                    } while (reader.NextResult());

                    // 2. Use the AsDataSet extension method
                    // var result = reader.AsDataSet();

                    // The result of each spreadsheet is in result.Tables
                }
            }

        }

        private static void sendEmail(string emailAddress)
        {

            var message = new MimeMessage();
            message.From.Add(new MailboxAddress(emailAddress));
            message.To.Add(new MailboxAddress(emailAddress));
            message.Subject = "Mail Kit Test";

            message.Body = new TextPart("plain")
            {
                Text = @"Hello!

                    This email is a test email. 

                    thanks, 

                    Jon"
            };

            using (var client = new SmtpClient())
            {
                // For demo-purposes, accept all SSL certificates (in case the server supports STARTTLS)
                client.ServerCertificateValidationCallback = (s, c, h, e) => true;
                Console.WriteLine("Connecting");
                client.Connect("smtp.sendgrid.net", 587, false);

            //Note: only needed if the SMTP server requires authentication
                
                client.Authenticate("username", "pasword");
                Console.WriteLine("Sending");
                client.Send(message);
                Console.WriteLine("Disconnecting");
                client.Disconnect(true);
            }

        }
    }
}
