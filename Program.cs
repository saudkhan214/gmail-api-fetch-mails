using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using System.Text.RegularExpressions;

namespace DeltaCRMWebMail
{
    class Program
    {

        static void Main(string[] args)
        {

            try
            {
                GetAllEmails();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex);
            }
        }

        public static void GetAllEmails()
        {
            try
            {
                string connectionString = ConfigurationManager.AppSettings.Get("connection");
                string format = "yyyy-MM-dd HH:mm:ss";
                Console.WriteLine("Coonecting with Gmail...\r\n");
                GmailService GmailService = GmailHelper.GetService();
                //List<Gmail> EmailList = new List<Gmail>();
                Console.WriteLine("Retriving emails...\r\n");

                UsersResource.MessagesResource.ListRequest ListRequest = GmailService.Users.Messages.List("me");
                ListRequest.LabelIds = "INBOX";
                ListRequest.IncludeSpamTrash = false;
                ListRequest.Q = "is:unread is:inbox category:primary "; //ONLY FOR UNDREAD EMAIL'S...
                ListRequest.MaxResults = Convert.ToInt32(ConfigurationManager.AppSettings.Get("emailToRetrive"));

                //GET ALL EMAILS
                ListMessagesResponse ListResponse = ListRequest.Execute();

                if (ListResponse != null && ListResponse.Messages != null)
                {
                    Console.WriteLine("Total {0} unread email(s)\r\n", ListResponse.Messages.Count);
                    //LOOP THROUGH EACH EMAIL AND GET WHAT FIELDS I WANT
                    foreach (Message Msg in ListResponse.Messages)
                    {
                        
                        UsersResource.MessagesResource.GetRequest Message = GmailService.Users.Messages.Get("me", Msg.Id);
                        Console.WriteLine("\n-----------------NEW MAIL----------------------");

                        //MAKE ANOTHER REQUEST FOR THAT EMAIL ID...
                        Message MsgContent = Message.Execute();
                        
                        if (MsgContent != null)
                        {
                            string FromAddress = string.Empty;
                            string Date = string.Empty;
                            string InternetDate = string.Empty;
                            string Subject = string.Empty;
                            string MailBody = string.Empty;
                            string ReadableText = string.Empty;
                            string DeleveredTo = string.Empty;
                            string FromName = string.Empty;
                            string Content_Type = string.Empty;
                            string FileExtension = string.Empty;
                            string FileName = string.Empty;
                            string messageGuid = string.Empty;

                            //LOOP THROUGH THE HEADERS AND GET THE FIELDS WE NEED (SUBJECT, MAIL)
                            foreach (var MessageParts in MsgContent.Payload.Headers)
                            {
                                if (MessageParts.Name == "From")
                                {
                                    FromAddress = MessageParts.Value;
                                }
                                else if (MessageParts.Name == "Date")
                                {
                                    Date = MessageParts.Value;
                                }
                                else if (MessageParts.Name == "Subject")
                                {
                                    Subject = MessageParts.Value;
                                }
                                else if (MessageParts.Name == "Delivered-To")
                                {
                                    DeleveredTo = MessageParts.Value;
                                }
                                else if (MessageParts.Name == "Content-Type")
                                {
                                    Content_Type = MessageParts.Value;
                                }

                            }
                            //READ MAIL BODY
                            Console.WriteLine("Reading Mail Body");
                            
                            //GET USER ID USING FROM EMAIL ADDRESS-------------------------------------------------------
                            string[] RectifyFromAddress = FromAddress.Split(' ');
                            string FromAdd = RectifyFromAddress[RectifyFromAddress.Length - 1];

                            if (!string.IsNullOrEmpty(FromAdd))
                            {
                                FromAdd = FromAdd.Replace("<", string.Empty);
                                FromAdd = FromAdd.Replace(">", string.Empty);
                                FromName = Regex.Replace(FromAddress, @"\<.*?\>", "");
                            }

                            //READ MAIL BODY-------------------------------------------------------------------------------------
                            MailBody = string.Empty;
                            if (MsgContent.Payload.Parts == null && MsgContent.Payload.Body != null)
                            {
                                MailBody = MsgContent.Payload.Body.Data;
                            }
                            else
                            {
                                MailBody = GmailHelper.MsgNestedParts(MsgContent.Payload.Parts);
                            }

                            Console.WriteLine("Saving emails into databae\r\n");
                            System.Threading.Thread.Sleep(500);
                            if (!string.IsNullOrEmpty(MailBody))
                            {
                                messageGuid = GUID();
                                string inserQuery = $"insert into WebmailsInfoMessages (WebmailMessagesGUID,FromEmail,Subject,Body,Created,FromName,ToEmail,ReceviedDate" +
                                    $") values('{messageGuid}',@fromEmail,@Subject,'{MailBody}','{DateTimeOffset.FromUnixTimeMilliseconds((long)MsgContent.InternalDate).DateTime.ToString(format)}',@fromName,'{DeleveredTo}','{DateTimeOffset.FromUnixTimeMilliseconds((long)MsgContent.InternalDate).DateTime.ToString(format)}')";

                                using (SqlConnection con = new SqlConnection(connectionString))
                                {
                                    con.Open();
                                    using (SqlCommand cmd = new SqlCommand(inserQuery, con))
                                    {
                                        List<SqlParameter> param = new List<SqlParameter>()
                                        {
                                            new SqlParameter("@fromEmail",DbType.String){Value=FromAdd},
                                            new SqlParameter("@fromName",DbType.String){Value=FromName},
                                            new SqlParameter("@Subject",DbType.String){Value=Subject},
                                        };
                                        cmd.Parameters.AddRange(param.ToArray());
                                        cmd.ExecuteNonQuery();
                                    }
                                }

                            }
                            Console.WriteLine("Reading Mail Attachments");
                            List<dynamic> FileObj = GmailHelper.GetAttachments("me", Msg.Id, Convert.ToString(ConfigurationManager.AppSettings["filePath"]));
                            if (FileObj.Count() > 0)
                            {
                                Console.WriteLine($"Attachments are saving at: { Convert.ToString(ConfigurationManager.AppSettings["filePath"])}\n");
                                //save file into db
                                foreach(var f in FileObj)
                                {
                                    FileInfo info = new FileInfo(f.Path);
                                    FileExtension = info.Extension;
                                    FileName = info.Name;
                                    string insertQuery = $"insert into WebMailAttachments (AttachmentsGUID,Attachments_Data,Attachments_MIME,Attachments_Extension,Attachments_FileName," +
                                $"WebMailInfoMessageGUID,AttachmentPath) values('{GUID()}',@Data,'{Content_Type.Split(';')[0]}','{FileExtension}',@Name,'{messageGuid}','{f.Path}')";

                                    using (SqlConnection con = new SqlConnection(connectionString))
                                    {
                                        con.Open();
                                        using (SqlCommand cmd = new SqlCommand(insertQuery, con))
                                        {
                                            //var param = new SqlParameter("@Data", SqlDbType.Binary);
                                            List<SqlParameter> prm = new List<SqlParameter>()
                                    {
                                        new SqlParameter("@Data",SqlDbType.VarBinary){Value=f.Data},
                                        new SqlParameter("@Name",DbType.String){Value=FileName},
                                    };
                                            cmd.Parameters.AddRange(prm.ToArray());
                                            cmd.ExecuteNonQuery();
                                        }
                                    }
                                }
                                
                            }
                            else
                            {
                                Console.WriteLine("Mail has no attachments.");
                            }
                        }
                        //MESSAGE MARKS AS READ AFTER READING MESSAGE
                         GmailHelper.MsgMarkAsRead("me", Msg.Id);
                    }

                    Console.WriteLine("Completed!");
                    System.Threading.Thread.Sleep(30000);
                }
                else
                {
                    Console.WriteLine("No unread email found");
                    Console.WriteLine("Completed!");
                    System.Threading.Thread.Sleep(30000);
                }
                
                //return EmailList;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex);
            }
        }
        private static string GUID()
        {
            return Guid.NewGuid().ToString();
        }
    }
}
