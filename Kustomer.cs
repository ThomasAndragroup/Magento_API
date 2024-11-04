using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Data.Odbc;
using System.Web;
using Newtonsoft.Json.Linq;
using System.Net;
using System.IO;
using Newtonsoft.Json;
using Microsoft.Office.Interop.Outlook;


namespace MagentoProductAPI
{
    class Kustomer
    {

        private static string KUSTOMERCSRMAILBOX = ConfigurationManager.AppSettings["KUSTOMERCSRMAILBOX"];
        private static string KUSTOMERCSRMAILBOXMAIN = ConfigurationManager.AppSettings["KUSTOMERCSRMAILBOXMAIN"];
        private static string KUSTOMERCSRMAILBOXCC = ConfigurationManager.AppSettings["KUSTOMERCSRMAILBOX"];


        public static Boolean Process_Kustomer_Contactus_CSREmailOnly(int MagentoContactusid = 0, int MaxRowstoProcess = 0)
        {
            DataTable dt;
            String body;

            if (MagentoContactusid > 0)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT *, ISNULL(Conversationid, '') [Convoid], ISNULL(BCC, '') [bccx], ISNULL(CC, '') [ccx] FROM communications..Magento_Contactus WHERE id = " + MagentoContactusid.ToString() + " ORDER BY id");
            }
            else
            {
                if (MaxRowstoProcess > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT TOP " + MaxRowstoProcess.ToString() + " *, ISNULL(Conversationid, '') [Convoid], ISNULL(BCC, '') [bccx], ISNULL(CC, '') [ccx] FROM communications..Magento_Contactus WHERE emailSent = 0 ORDER BY id");
                }
                else
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT *, ISNULL(Conversationid, '') [Convoid], ISNULL(BCC, '') [bccx], ISNULL(CC, '') [ccx] FROM communications..Magento_Contactus WHERE emailSent = 0 ORDER BY id");
                }
            }

            foreach (DataRow dr in dt.Rows)
            {
                if (MagetnoProductAPI.DevMode == 1)
                {
                    Console.WriteLine(dr["kustomeremail"].ToString() + "; " + dr["kustomername"].ToString() + "; " + dr["id"].ToString());
                }

                body = dr["emailBody"].ToString();
                body = body.Replace("'", "''");

                // EMAIL TO CSR 
                //Helper.SendEmail(dr["emailSubject"].ToString(), body, dr["emailTo"].ToString(), dr["emailCustomer"].ToString(), dr["ccx"].ToString(), dr["bccx"].ToString());
                //2024-04-18 Update to send BCCs ONLY, not to CSR mailboxes 
                //Helper.QuickEmail(dr["emailSubject"].ToString(), body, dr["emailTo"].ToString(), dr["emailCustomer"].ToString(), dr["ccx"].ToString(), dr["bccx"].ToString());
                Helper.QuickEmail(dr["emailSubject"].ToString(), body, KUSTOMERCSRMAILBOX, dr["emailCustomer"].ToString(), dr["ccx"].ToString(), dr["bccx"].ToString());
                
                Helper.Sql_Misc_NonQuery("UPDATE communications..Magento_Contactus SET emailSent=3 WHERE ID = " + dr["id"].ToString());

            }
            return true;
        }

        // NEED THIS? 
        //2024-10-18 - New FOR API Get of contactus, questions
        //ProcessType: question || contactus
        //Contactus rows only need to get pushed to CSR mailbox, they get pushed to Kustomer some other way
        //PDP Questions are not pushed to Kustomer so they are pused to CSR mailbox also
        public static Boolean Process_Kustomer_Contactus_Questions(string ProcessType = "", int MaxRowstoProcess = 20)
        {
            Boolean ReturnValue = true;
            DataTable dt;

            if (ProcessType.ToLower() == "question" || ProcessType.ToLower() == "questions")
            {
                dt = Helper.Sql_Misc_Fetch("SELECT TOP " + MaxRowstoProcess + " id FROM communications..Magento_Contactus "
                    + " WHERE emailSent IN (0, 1) AND LEN(ISNULL(kustomeremail, '')) > 0 AND ISNULL(conversationid,'') = '' AND Datestamp > DateAdd(week, -1, Getdate()) "
                    + " AND ISNULL(Magentoid_Question,0) > 0 ORDER BY id ");

                foreach (DataRow dr in dt.Rows)
                {
                    Kustomer.Process_Kustomer_Contactus(int.Parse(dr["id"].ToString()), 1, false, true);
                }
            }
            else
            {
                dt = Helper.Sql_Misc_Fetch("SELECT TOP " + MaxRowstoProcess + " id FROM communications..Magento_Contactus "
                   + " WHERE emailSent IN (0, 1) AND LEN(ISNULL(kustomeremail, '')) > 0 AND ISNULL(conversationid,'') = '' AND Datestamp > DateAdd(week, -1, Getdate()) "
                   + " AND ISNULL(Magentoid_Contactus,0) > 0 ORDER BY id ");

                foreach (DataRow dr in dt.Rows)
                {
                    Kustomer.Process_Kustomer_Contactus(int.Parse(dr["id"].ToString()), 1, true, false);
                }
            }

            return ReturnValue;
        }

        public static Boolean Process_Kustomer_Contactus(int MagentoContactusid = 0, int MaxRowstoProcess = 0, Boolean SendCSREmail = true, Boolean SendAPI = true)
        {
            Boolean ReturnValue = true;
            DataTable dt;
            String Conversationid = "";
            String CSRBcc = "Customer Service<customerservice@herroom.com>;";

            if (MagentoContactusid > 0)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT *, ISNULL(Conversationid, '') [Convoid], ISNULL(BCC, '') [bccx], ISNULL(CC, '') [ccx] FROM communications..Magento_Contactus WHERE id = " + MagentoContactusid.ToString() + " ORDER BY id");
            }
            else if (!SendAPI)
            {
                if (MaxRowstoProcess == 0)
                {
                    MaxRowstoProcess = 2;
                }

                //dt = Helper.Sql_Misc_Fetch("SELECT TOP " + MaxRowstoProcess + " *, ISNULL(Conversationid, '') [Convoid], ISNULL(BCC, '') [bccx], ISNULL(CC, '') [ccx] "
                //    + " FROM communications..Magento_Contactus "
                //    + " WHERE emailSent IN (0, 1) AND LEN(ISNULL(kustomeremail, '')) > 0 AND ISNULL(conversationid,'') = '' AND Datestamp > DateAdd(week, -1, Getdate())  ORDER BY id ");
                dt = Helper.Sql_Misc_Fetch("SELECT TOP " + MaxRowstoProcess + " *, ISNULL(Conversationid, '') [Convoid], ISNULL(BCC, '') [bccx], ISNULL(CC, '') [ccx] "
                    + " FROM communications..Magento_Contactus "
                    + " WHERE emailSent IN (0, 1) AND LEN(ISNULL(kustomeremail, '')) > 0 AND Datestamp > DateAdd(week, -1, Getdate())  ORDER BY id ");
            }
            else
            {
                //2024-04-22 Changed WHERE EmailSent from (0,100), changing AskaQuestion sproc to set emailsent to 1 
                if (MaxRowstoProcess > 0)
                {
                    //dt = Helper.Sql_Misc_Fetch("SELECT TOP " + MaxRowstoProcess.ToString() + " *, ISNULL(Conversationid, '') [Convoid], ISNULL(BCC, '') [bccx], ISNULL(CC, '') [ccx] FROM communications..Magento_Contactus WHERE emailSent IN (1, 100) AND LEN(ISNULL(kustomeremail, '')) > 0 ORDER BY id");
                    dt = Helper.Sql_Misc_Fetch("SELECT TOP " + MaxRowstoProcess.ToString() + " *, ISNULL(Conversationid, '') [Convoid], ISNULL(BCC, '') [bccx], ISNULL(CC, '') [ccx] FROM communications..Magento_Contactus WHERE emailSent IN (0,1) AND LEN(ISNULL(kustomeremail, '')) > 0 ORDER BY id");
                }
                else
                {
                    //dt = Helper.Sql_Misc_Fetch("SELECT *, ISNULL(Conversationid, '') [Convoid], ISNULL(BCC, '') [bccx], ISNULL(CC, '') [ccx] FROM communications..Magento_Contactus WHERE emailSent IN (1,100) AND LEN(ISNULL(kustomeremail, '')) > 0 ORDER BY id");
                    dt = Helper.Sql_Misc_Fetch("SELECT *, ISNULL(Conversationid, '') [Convoid], ISNULL(BCC, '') [bccx], ISNULL(CC, '') [ccx] FROM communications..Magento_Contactus WHERE emailSent IN (0,1) AND LEN(ISNULL(kustomeremail, '')) > 0 ORDER BY id");
                }
            }

            foreach (DataRow dr in dt.Rows)
            {
                if (MagetnoProductAPI.DevMode == 1)
                {
                    Console.WriteLine(dr["kustomeremail"].ToString() + "; " + dr["kustomername"].ToString() + "; " + dr["id"].ToString());
                }

                if (SendAPI)
                {
                    Conversationid = KustomerApiPush(dr["kustomeremail"].ToString(), dr["kustomername"].ToString(), dr["emailsubject"].ToString(), dr["emailbody"].ToString().Replace("'","`"), "").ToUpper();
                }
                else
                {
                    Conversationid = "98";
                }

                if (Conversationid != "ERROR")
                {
                    Helper.Sql_Misc_NonQuery("UPDATE communications..Magento_Contactus SET emailSent=200, Conversationid = '" + Conversationid + "' WHERE ID = " + dr["id"].ToString() + " AND LEN(conversationid) < 4");

                    // EMAIL TO CSR 
                    if (SendCSREmail)
                    {
                        if (dr["emailTo"].ToString().ToLower() == "customerservice@hisroom.com")
                        {
                            CSRBcc = "Customer Service<customerservice@hisroom.com>;";
                        }

                        if (MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine("CSRBCC: " + CSRBcc + dr["bccx"].ToString());
                            Console.WriteLine(" ------------------------ ");
                            Console.WriteLine(@"Helper.QuickEmail(""" + dr["emailSubject"].ToString().Replace("'", "`") + @""", """ + dr["emailBody"].ToString().Replace("'", "`") + @""", """ + KUSTOMERCSRMAILBOX + @""" +, @""" + dr["emailCustomer"].ToString().Replace("'", "") + @""", """", """ + CSRBcc + dr["bccx"].ToString() + @""");");
                        }

                        //Changed email for functions for more reliability
                        //Helper.SendEmail(dr["emailSubject"].ToString(), dr["emailBody"].ToString(), KUSTOMERCSRMAILBOX, dr["emailCustomer"].ToString(), "", CSRBcc + dr["bccx"].ToString());

                        //2024-08-20 Changed main email to be 'from' customerservice@herroom.com to these emails will go out from Kustomer correctly
                        //Helper.QuickEmail(dr["emailSubject"].ToString(), dr["emailBody"].ToString(), KUSTOMERCSRMAILBOX, dr["emailCustomer"].ToString(), "", CSRBcc + dr["bccx"].ToString());
                        //Helper.QuickEmail(dr["emailSubject"].ToString(), dr["emailBody"].ToString(), KUSTOMERCSRMAILBOXMAIN, dr["emailCustomer"].ToString(), KUSTOMERCSRMAILBOX, CSRBcc + dr["bccx"].ToString());
                        //Helper.QuickEmail(dr["emailSubject"].ToString(), dr["emailBody"].ToString(), dr["emailTo"].ToString(), dr["emailCustomer"].ToString(), KUSTOMERCSRMAILBOX, CSRBcc + dr["bccx"].ToString());

                        //2024-10-18 DUplicate Kustomer issue (API Get Change)
                        //Helper.QuickEmail(dr["emailSubject"].ToString(), dr["emailBody"].ToString(), dr["emailTo"].ToString(), dr["emailCustomer"].ToString(), "",  KUSTOMERCSRMAILBOX +";" + CSRBcc + dr["bccx"].ToString());
                        Helper.QuickEmail(dr["emailSubject"].ToString(), dr["emailBody"].ToString().Replace("'", "`"), dr["emailTo"].ToString(), dr["emailCustomer"].ToString().Replace("'", ""), "", dr["bccx"].ToString());

                        Helper.Sql_Misc_NonQuery("UPDATE communications..Magento_Contactus SET emailSent=200 WHERE ID = " + dr["id"].ToString());

                    }
                }
                else
                {
                    Helper.Sql_Misc_NonQuery("UPDATE communications..Magento_Contactus SET emailSent=-2 WHERE ID = " + dr["id"].ToString());

                    // EMAIL TO CSR 
                    Helper.SendEmail(dr["emailSubject"].ToString(), "NOT IN KUSTOMER, API ERROR :: " + dr["emailBody"].ToString().Replace("'", "`"), dr["emailTo"].ToString(), dr["emailCustomer"].ToString(), dr["ccx"].ToString(), dr["bccx"].ToString());
                } 

                System.Threading.Thread.Sleep(1500);
            }

            return ReturnValue;
        }

        private static String KustomerApiPush(String Email, String KustomerName, String Subject, String Message, String ConversationId)
        {
            String Returnvalue = "";
            String json;
            DataRow dr;
            dynamic dJson;
            string EID = "0";
            string EndpointName = "";
            DataTable dt;
            string Kustomerid = "";
            String ConvoId = "";
            int KustomerCustId = 0;

            /*
             This needs to: 
                 1) fetch kustomerid of email
                 2) if "", then send API customer created and snag Id && update Kustomer_Cust table with
                 3) API Create a conversation, get ID, update kustomer_data table with kustomercustid, transactionid
                 4) Create a message for convo, get ID and save it
             */

            //Will always return a now
            dt = Helper.Sql_Misc_Fetch("SELECT TOP 1 REPLACE(email,'@','%40') [email], kustomerid, id FROM( "
                + "SELECT email, kustomerid, id, 0[ord] from Kustomer_Cust WHERE email = '" + Email + "' "
                + "UNION SELECT '', '', 0, 1) ZZZ ORDER BY[ord] ");

            dr = dt.Rows[0];
            if (dr["kustomerid"].ToString() == "")
            {
                // Find Existing Email
                EndpointName = "https://api.kustomerapp.com/v1/customers/email=" + Email.Replace("@", "%40");
                Console.WriteLine(EndpointName);

                var httpWebRequestEmail = (HttpWebRequest)WebRequest.Create(EndpointName);
                httpWebRequestEmail.ContentType = "application/json";
                httpWebRequestEmail.KeepAlive = true;
                httpWebRequestEmail.UserAgent = "PostmanRuntime/7.31.1";
                httpWebRequestEmail.Timeout = 60000;
                httpWebRequestEmail.Headers.Add("Authorization", "Bearer " + ConfigurationManager.AppSettings["BearerTokenKustomer"]);
                httpWebRequestEmail.Method = "GET";
                try
                {
                    var httpResponseEmail = (HttpWebResponse)httpWebRequestEmail.GetResponse();
                    var streamReaderEmail = new StreamReader(httpResponseEmail.GetResponseStream());
                    var resultEmail = streamReaderEmail.ReadToEnd();

                    dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(resultEmail);
                    EID = dJson.data.id;
                    Console.WriteLine("Custid: " + EID.ToString());
                    Kustomerid = EID.ToString();
                }
                catch (System.Exception exEmail)
                {
                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("exEmail: " + exEmail.ToString());
                    }
                    Returnvalue = "ERROR";
                    Kustomerid = "";
                }

                try
                {
                    if (Kustomerid != "")
                    {
                        Helper.Sql_Misc_NonQuery("INSERT Communications..Kustomer_cust(email, Kustomerid) SELECT '" + Email + "', '" + Kustomerid + "'");

                        DataTable dt2 = Helper.Sql_Misc_Fetch("SELECT id FROM Communications..Kustomer_cust where email = '" + Email + "';");
                        if (dt2.Rows.Count == 1)
                        {
                            KustomerCustId = int.Parse(dt2.Rows[0]["id"].ToString());
                        }
                    }
                    else
                    {
                        Kustomerid = CreateKustomer(Email, KustomerName);

                        DataTable dtID;
                        dtID = Helper.Sql_Misc_Fetch("SELECT TOP 1 id FROM Communications..Kustomer_cust WHERE Email = '" + Email.ToLower() + "';");
                        if (dtID.Rows.Count == 1)
                        {
                            KustomerCustId = int.Parse(dtID.Rows[0]["id"].ToString());
                        }

                        //Helper.Sql_Misc_NonQuery("INSERT ")
                    }
                }
                catch (System.Exception exEmail2)
                {
                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("exEmail2: " + exEmail2.ToString());
                    }
                    Returnvalue = "ERROR";
                    Kustomerid = "";
                }
            }
            else
            {
                Kustomerid = dt.Rows[0]["kustomerid"].ToString();
                KustomerCustId = int.Parse(dt.Rows[0]["id"].ToString());
            }


            if (Kustomerid != "")
            {
                /////////////////////////////////////////////////////////////////////////////////
                //CONVERSATION TEST 
                EndpointName = "https://api.kustomerapp.com/v1/conversations";
                var httpWebRequestConvo = (HttpWebRequest)WebRequest.Create(EndpointName);

                httpWebRequestConvo.ContentType = "application/json";
                httpWebRequestConvo.KeepAlive = true;
                httpWebRequestConvo.UserAgent = "PostmanRuntime/7.31.1";
                httpWebRequestConvo.Timeout = 60000;
                httpWebRequestConvo.Headers.Add("Authorization", "Bearer " + ConfigurationManager.AppSettings["BearerTokenKustomer"]);
                httpWebRequestConvo.Method = "POST";

                json = @"{ ""customer"": """ + Kustomerid + @""", ""name"": """ + Subject + @""", ""priority"": 2, ""status"": ""open""}";

                try
                {
                    StreamWriter swNote = new StreamWriter(httpWebRequestConvo.GetRequestStream());
                    swNote.Write(json);
                    swNote.Close();
                    swNote.Dispose();

                    var httpResponseConvo = (HttpWebResponse)httpWebRequestConvo.GetResponse();

                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine(httpResponseConvo.StatusCode.ToString().ToString().ToLower());
                    }

                    if (httpResponseConvo.StatusCode.ToString().ToString().ToLower() == "created" || httpResponseConvo.StatusCode.ToString().ToString().ToLower() == "ok")
                    {
                        var streamReader = new StreamReader(httpResponseConvo.GetResponseStream());
                        var result = streamReader.ReadToEnd();

                        dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(result);
                        EID = dJson.data.id;
                        if (MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine("EID: " + EID.ToString());
                        }
                        ConvoId = EID.ToString();
                        Returnvalue = ConvoId; 
                    }
                    else
                    {
                        Console.WriteLine("ERROR1x: " + httpResponseConvo.StatusCode.ToString().ToString().ToLower());
                        Returnvalue = "ERROR";
                    }

                }
                catch (System.Exception ex1)
                {
                    Console.WriteLine("CATCH 1 ::  " + ex1.ToString());

                    Returnvalue = "ERROR";
                }
            }

            if (MagetnoProductAPI.DevMode == 1)
            {
                Console.WriteLine(Message);
            }

            //ADD NOTE, REPLACE WITH MESSAGE ASAP 
            if (!AddConversationMessage(Kustomerid, ConvoId, Message, Email, ""))
            {
                Returnvalue = "ERROR";
            }

            if (Returnvalue.Length == 0)
            {
                Helper.Sql_Misc_NonQuery("INSERT Communications..Kustomer_Data(KustomerCustId, TransactionType, Conversationid, Datestamp) SELECT " + KustomerCustId.ToString() + ", 'Conversation', '" + ConvoId + "', Getdate()");
            }

            return Returnvalue;
        }

        public static Boolean AddConversationNote(String Conversationid, String Message)
        {
            Boolean Returnvalue = true;
            String EndpointName;
            String json;

            if (Conversationid.Length > 0)
            {
                //////////////////////////////////////////////////////////////////////////////////////////
                //NOTES WORKS
                //Create Note to begin with
                EndpointName = "https://api.kustomerapp.com/v1/conversations/" + Conversationid + "/notes";

                var httpWebRequestNote = (HttpWebRequest)WebRequest.Create(EndpointName);

                httpWebRequestNote.ContentType = "application/json";
                httpWebRequestNote.KeepAlive = true;
                httpWebRequestNote.UserAgent = "PostmanRuntime/7.31.1";
                httpWebRequestNote.Timeout = 60000;
                httpWebRequestNote.Headers.Add("Authorization", "Bearer " + ConfigurationManager.AppSettings["BearerTokenKustomer"]);
                httpWebRequestNote.Method = "POST";

                try
                {
                    Message = Message.Replace("\"", "'");
                    Message = Message.Replace("<br />", "; ");
                    Message = Message.Replace("<br>", "; ");
                    Message = Message.Replace("<p>", "; ");
                    Message = Message.Replace("<p>", "; ");
                    Message = Message.Replace("\r\n", "; ");

                    json = @"{ ""body"": """ + Message + @""" }";

                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine(json);
                    }

                    StreamWriter swNote = new StreamWriter(httpWebRequestNote.GetRequestStream());
                    swNote.Write(json);
                    swNote.Close();
                    swNote.Dispose();

                    var httpResponseNote = (HttpWebResponse)httpWebRequestNote.GetResponse();
                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine(httpResponseNote.StatusCode.ToString().ToString().ToLower());
                    }

                    if (httpResponseNote.StatusCode.ToString().ToString().ToLower() == "created" || httpResponseNote.StatusCode.ToString().ToString().ToLower() == "ok")
                    {
                        //NOTES WILL NOT HAVE ID - do nothing for now
                    }
                    else
                    {
                        Console.WriteLine("ERROR: " + httpResponseNote.StatusCode.ToString().ToString().ToLower());
                        Returnvalue = false;
                    }
                }
                catch (System.Exception ex)
                {
                    Console.WriteLine("CATCH::  " + ex.ToString());
                    Returnvalue = false;
                }
            }

            return Returnvalue;
        }

        private static String CreateKustomer(String Email, String KustomerName)
        {
            String Kustomerid = "";
            string json;
            dynamic dJson;
            string EID = "";

            var httpWebRequest = (HttpWebRequest)WebRequest.Create("https://api.kustomerapp.com/v1/customers");

            httpWebRequest.ContentType = "application/json";
            httpWebRequest.KeepAlive = true;
            httpWebRequest.UserAgent = "PostmanRuntime/7.31.1";
            httpWebRequest.Timeout = 60000;
            httpWebRequest.Headers.Add("Authorization", "Bearer " + ConfigurationManager.AppSettings["BearerTokenKustomer"]);
            httpWebRequest.Method = "POST";

            json = @"{""sentiment"": {""polarity"": 0, ""confidence"": 0}, "
                + @" ""name"": """ + KustomerName + @""",""company"": """", ""emails"": [ "
                + @" { ""email"": """ + Email.ToLower() + @""",""verified"": true, ""type"": ""other""}]}";

            if (MagetnoProductAPI.DevMode > 0)
            {
                Console.WriteLine(json);
            }

            try
            {
                StreamWriter sw = new StreamWriter(httpWebRequest.GetRequestStream());
                sw.Write(json);
                sw.Close();
                sw.Dispose();

                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();

                if (httpResponse.StatusCode.ToString().ToString().ToLower() == "created" || httpResponse.StatusCode.ToString().ToString().ToLower() == "ok")
                {
                    var streamReader = new StreamReader(httpResponse.GetResponseStream());
                    var result = streamReader.ReadToEnd();

                    dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(result);
                    Console.WriteLine("dJson: " + dJson.ToString());

                    EID = dJson.data.id;
                    Console.WriteLine("EID: " + EID.ToString());
                    Kustomerid = EID.ToString();
                }
                else
                {
                    Console.WriteLine("ERROR: " + httpResponse.StatusCode.ToString().ToString().ToLower());
                    Kustomerid = "ERROR";
                }

            }
            catch (System.Exception ex1)
            {
                Console.WriteLine("CATCH 1 ::  " + ex1.ToString());

                Kustomerid = "ERROR";
            }

            if (Kustomerid != "" && Kustomerid != "ERROR")
            {
                Helper.Sql_Misc_NonQuery("INSERT Communications..Kustomer_cust(email, Kustomerid) SELECT '" + Email.ToLower() + "', '" + Kustomerid + "'");
            }

            return Kustomerid;
        }

        private static Boolean AddConversationMessage(string Customerid, string Conversationid, string Message, string FromEmail, string ToEmail)
        {
            String json;

            //Message = Helper.HtmlToPlainText(Message);  //Line Breaks Confuse JSON
            Message = Message.Replace("\"", "'");
            Message = Message.Replace("<br />", @"\r");   
            Message = Message.Replace("<br>", @"\r");
            Message = Message.Replace("<p>", @"\r");
            Message = Message.Replace("\r\n", @"\r");

            if (ToEmail == "")
            {
                if (Message.ToLower().Contains("website: hisroom") || Message.ToLower().Contains("website:hisroom"))
                {
                    ToEmail = "customerservice@hisroom.com";
                }
                else
                {
                    ToEmail = "customerservice@herroom.com";
                }
            }

            json = @"{""app"": ""postmark"",""channel"": ""email"", ""customer"": """ + Customerid + @""",""conversation"": """ + Conversationid + @""",""direction"": ""in"",""preview"": """ + Message + @""",""meta"": {""from"": """ + FromEmail + @""",""to"": [{""email"": """ + ToEmail + @"""}]}}";

            if (MagetnoProductAPI.DevMode > 0)
            {
                Console.WriteLine(" ------------------------ ");
                Console.WriteLine(json);
            }

            var httpWebRequest = (HttpWebRequest)WebRequest.Create("https://api.kustomerapp.com/v1/messages");

            httpWebRequest.ContentType = "application/json";
            httpWebRequest.KeepAlive = true;
            httpWebRequest.UserAgent = "PostmanRuntime/7.31.1";
            httpWebRequest.Timeout = 60000;
            httpWebRequest.Headers.Add("Authorization", "Bearer " + ConfigurationManager.AppSettings["BearerTokenKustomer"]);
            httpWebRequest.Method = "POST";

            try
            {
                StreamWriter sw = new StreamWriter(httpWebRequest.GetRequestStream());
                sw.Write(json);
                sw.Close();
                sw.Dispose();

                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                Console.WriteLine(httpResponse.StatusCode.ToString().ToString().ToLower());

            }
            catch (System.Exception ex)
            {
                Console.WriteLine("ERROR: " + ex.ToString());
                return false;
            }
            return true;
        }

        public static Boolean Outlook_Reader()
        {
            String Convoid = "";
            Application oApp = new Application();
            NameSpace oNS = (NameSpace)oApp.GetNamespace("MAPI");

            MAPIFolder theInbox = oNS.Folders["thomas@andragroup.com"].Folders["Inbox"];

            Items unreadItems = theInbox.Items.Restrict("[Unread]=true");
            string t = "thomas.tribble@gmail.com";
            try
            {
                foreach (MailItem it in unreadItems)
                {
                   

                    if (it.SenderEmailAddress == t)        //t.Subject
                    {
                        MailItem replyMail = it.Reply();

                        replyMail.HTMLBody = "REPLYING";

                        if (it.Body.ToString().Contains("CID:"))
                        {
                            //Harvest CID:
                            Convoid = it.Body.ToUpper().Substring(it.Body.IndexOf("CID:") + 4);
                            Convoid = Convoid.Substring(0, Convoid.IndexOf(";"));
                            Console.WriteLine("IT BODY: " + it.Body);
                            Console.WriteLine();
                            Console.WriteLine("Convoid: " + Convoid);
                        }

                        Console.WriteLine("Subject: " + it.Subject);
                        Console.WriteLine("Sender: " + it.SenderEmailAddress);
                        Console.WriteLine();
      
                    }
                }
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            return true;
        }

        public static Boolean TestforContactusinKustomer(long MagentoContactusid)
        {
            String EndpointName;
            dynamic dJson;
            string Preview = "";
            DataTable dt;
            //string EmailBody;
            //string hEmailBody;
            string Kustomerid = "";
            string EID = "0";
            Boolean bReturnvalue = true;
            String convoid = "";
            Boolean Found = false;
            string kustomeremail;
            string formType;
            string orderNumber;
            string comment255;
            string datestamp;
            string pageName;
            string phone1;
            string DBKustomerid = "";
            string Rowid; 

            //2024-03-05 - Changed the DB call to return cust..phone1 and also use EmailTo instead of KustomerEmail if begin sent o customerservice mailbox... TE wants those going to customers' queue when possible
            //dt = Helper.Sql_Misc_Fetch("SELECT ISNULL(kustomeremail, 'xxx') [kustomeremail], ISNULL(Ordernumber,'xxx') [ordernumber], ISNULL(formType, 'xxx') [formtype], REPLACE(REPLACE(REPLACE(ISNULL(Comment255,'xxx'), ' ', ''), char(10), ''), char(13) , '') [comment255], Datestamp, REPLACE(REPLACE(REPLACE(REPLACE(ISNULL(CC.Phone1,''),'-', ''), '+', ''), 'US', ''), '9729796896', '') [phone] FROM Communications..Magento_Contactus MC with (NOLOCK) LEFT OUTER JOIN HerCust..Cust CC with (NOLOCK) ON CC.Email = MC.KustomerEmail WHERE ID = " + MagentoContactusid.ToString());

            dt = Helper.Sql_Misc_Fetch("SELECT ISNULL(CASE WHEN SUBSTRING(MC.KustomerEmail, 0, 16) ='customerservice' THEN EmailTo ELSE mc.KustomerEmail END, 'xxx') [kustomeremail] " +
                " , ISNULL(Ordernumber, 'xxx')[ordernumber] " +
                " , ISNULL(formType, 'xxx')[formtype], REPLACE(REPLACE(REPLACE(ISNULL(Comment255, 'xxx'), ' ', '') " +
                " , char(10), ''), char(13), '')[comment255], Datestamp " +
                " , REPLACE(REPLACE(REPLACE(REPLACE(ISNULL(CC.Phone1, ''), '-', ''), '+', ''), 'US', ''), '9729796896', '') [Phone] " +
                " , phone1, SUBSTRING(MC.KustomerEmail, 0, 16), ISNULL(KC.kustomerid,'') [kustomerid], MC.id [rowid]  " +
                " FROM Communications..Magento_Contactus MC " +
                " LEFT OUTER JOIN HerCust..Cust CC ON CC.Email = CASE WHEN SUBSTRING(MC.KustomerEmail, 0, 16) = 'customerservice' THEN EmailTo ELSE mc.KustomerEmail END " +
                " LEFT OUTER JOIN Communications..Kustomer_Cust KC ON KC.email = MC.KustomerEmail " +
                " WHERE MC.ID = " + MagentoContactusid.ToString());

            if (dt.Rows.Count == 1)
            {
                kustomeremail = dt.Rows[0]["kustomeremail"].ToString();
                formType = dt.Rows[0]["formtype"].ToString();
                orderNumber = dt.Rows[0]["orderNumber"].ToString();
                comment255 = dt.Rows[0]["comment255"].ToString();
                datestamp = dt.Rows[0]["datestamp"].ToString();
                phone1 = dt.Rows[0]["phone"].ToString();
                DBKustomerid = dt.Rows[0]["kustomerid"].ToString();
                Rowid = dt.Rows[0]["rowid"].ToString();

                if (comment255.Length > 100)
                {
                    comment255 = comment255.Substring(0, 100);
                }
                comment255 = comment255.Replace("'", "''");

                if (DBKustomerid == "")
                {
                    // Find Existing Email
                    EndpointName = "https://api.kustomerapp.com/v1/customers/email=" + kustomeremail.Replace("@", "%40");
                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine(EndpointName);
                    }

                    var httpWebRequestEmail = (HttpWebRequest)WebRequest.Create(EndpointName);
                    httpWebRequestEmail.ContentType = "application/json";
                    httpWebRequestEmail.KeepAlive = true;
                    httpWebRequestEmail.UserAgent = "PostmanRuntime/7.31.1";
                    httpWebRequestEmail.Timeout = 60000;
                    httpWebRequestEmail.Headers.Add("Authorization", "Bearer " + ConfigurationManager.AppSettings["BearerTokenKustomer"]);
                    httpWebRequestEmail.Method = "GET";
                    try
                    {
                        var httpResponseEmail = (HttpWebResponse)httpWebRequestEmail.GetResponse();
                        var streamReaderEmail = new StreamReader(httpResponseEmail.GetResponseStream());
                        var resultEmail = streamReaderEmail.ReadToEnd();

                        dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(resultEmail);
                        EID = dJson.data.id;
                        Console.WriteLine("Custid: " + EID.ToString());
                        Kustomerid = EID.ToString();
                        if (MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine("Kustomerid = " + Kustomerid);
                        }
                    }
                    catch (System.Exception exEmail)
                    {
                        if (MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine("exEmail: " + exEmail.ToString());
                        }
                        bReturnvalue = false;
                        Kustomerid = "";
                    }
                }
                else
                {
                    Kustomerid = DBKustomerid; 
                }

                //2024-03-05 LOOK FOR CUSTOMER BY PHONE ALSO
                if (Kustomerid == "" && phone1 != "")
                {
                    // Find Existing Email
                    EndpointName = "https://api.kustomerapp.com/v1/customers/phone=" + phone1 ;

                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine(EndpointName);
                    }

                    var httpWebRequestPhone = (HttpWebRequest)WebRequest.Create(EndpointName);
                    httpWebRequestPhone.ContentType = "application/json";
                    httpWebRequestPhone.KeepAlive = true;
                    httpWebRequestPhone.UserAgent = "PostmanRuntime/7.31.1";
                    httpWebRequestPhone.Timeout = 60000;
                    httpWebRequestPhone.Headers.Add("Authorization", "Bearer " + ConfigurationManager.AppSettings["BearerTokenKustomer"]);
                    httpWebRequestPhone.Method = "GET";
                    try
                    {
                        var httpResponsePhone = (HttpWebResponse)httpWebRequestPhone.GetResponse();
                        var streamReaderPhone = new StreamReader(httpResponsePhone.GetResponseStream());
                        var resultPhone = streamReaderPhone.ReadToEnd();

                        dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(resultPhone);
                        EID = dJson.data.id;
                        Console.WriteLine("Custid: " + EID.ToString());
                        Kustomerid = EID.ToString();
                        if (MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine("Kustomerid-2 = " + Kustomerid);
                        }
                    }
                    catch (System.Exception exPhone)
                    {
                        if (MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine("exPhone: " + exPhone);
                        }
                        bReturnvalue = false;
                        Kustomerid = "";
                    }
                }

                //Kustomer-Customer found
                if (Kustomerid == "")
                {
                    Helper.Sql_Misc_NonQuery("UPDATE Communications..magento_contactus SET Conversationid = '-2', emailsent = 100 WHERE ID = " + MagentoContactusid.ToString());
                    Helper.SendEmail("Magento_ContactUS not in Kustomer", "Magento_ContactUS !CUSTOMER! NOT in Kustomer: ID: " + MagentoContactusid.ToString() + "; " + kustomeremail + "; TYPE: " + formType.ToString() + "; Date: " + datestamp.ToString(), "thomas@andragroup.com", "alerts@andragroup.com");

                }
                else //if (Kustomerid != "")
                {
                    if (DBKustomerid == "")
                    {
                        Helper.Sql_Misc_NonQuery(" IF 0 = (SELECT COUNT(*) FROM Communications..Kustomer_Cust WHERE email = 'mreiner33@gmail.com') "
                            + " INSERT Communications..Kustomer_Cust(email, KustomerId) SELECT '" + kustomeremail + "', '" + Kustomerid + "'; "
                            + " ELSE UPDATE Communications..Kustomer_Cust SET KustomerId = '" + Kustomerid + "' WHERE email = '" + kustomeremail + "' AND ISNULL(KustomerId, '') = ''; ");
                    }

                    EndpointName = "https://api.kustomerapp.com/v1/customers/" + Kustomerid + "/conversations?page=1&pageSize=100&fromDate=" + datestamp;
                    Console.WriteLine(EndpointName);

                    var httpWebRequestConvo = (HttpWebRequest)WebRequest.Create(EndpointName);
                    httpWebRequestConvo.ContentType = "application/json";
                    httpWebRequestConvo.KeepAlive = true;
                    httpWebRequestConvo.UserAgent = "PostmanRuntime/7.31.1";
                    httpWebRequestConvo.Timeout = 60000;
                    httpWebRequestConvo.Headers.Add("Authorization", "Bearer " + ConfigurationManager.AppSettings["BearerTokenKustomer"]);
                    httpWebRequestConvo.Method = "GET";
                    try
                    {
                        var httpResponseEmail = (HttpWebResponse)httpWebRequestConvo.GetResponse();
                        var streamReaderEmail = new StreamReader(httpResponseEmail.GetResponseStream());
                        var resultEmail = streamReaderEmail.ReadToEnd();

                        dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(resultEmail);
                        if (MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine(dJson);
                            Console.WriteLine(dJson.data.Count.ToString());
                        }

                        try
                        {
                            for (int xx = 0; xx < dJson.data.Count; xx++)
                            {
                                Preview = dJson.data[xx].attributes.preview;                 
                                pageName = dJson.data[xx].attributes.name;
                                convoid = dJson.data[xx].id;

                                Preview = Preview.Replace("\n", "");
                                Preview = Preview.Replace(" ", "");
                                comment255 = comment255.Replace("\t", "");

                                if (MagetnoProductAPI.DevMode > 0)
                                {
                                    Console.WriteLine(kustomeremail + "; " + formType);
                                    Console.WriteLine(Preview.Replace("\n", " "));
                                    Console.WriteLine("Comment255: " + comment255);
                                    Console.WriteLine("------------------------");
                                    Console.WriteLine(Preview.Contains(kustomeremail));
                                    //Console.WriteLine(pageName.Contains(formType));
                                    Console.WriteLine(Preview.ToLower().Contains("subject" + formType.ToLower().Replace(" ", "")));
                                    Console.WriteLine(Preview.Replace("\n", " ").Contains(comment255));
                                    Console.WriteLine("------------------------");
                                }

                             //   if (Preview.Contains(kustomeremail) && pageName.Contains(formType) && Preview.Replace("\n", " ").Contains(comment255))
                                 if (Preview.Contains(kustomeremail) && Preview.ToLower().Contains("subject" + formType.ToLower().Replace(" ", "")) && Preview.Replace("\n", " ").Contains(comment255))
                                    {
                                    if (MagetnoProductAPI.DevMode > 0)
                                    {
                                        Console.WriteLine("Match");
                                    }
                                    else
                                    {
                                        Helper.Sql_Misc_NonQuery("UPDATE Communications..magento_contactus set Conversationid = '" + convoid + "' WHERE ID = " + MagentoContactusid.ToString());
                                    }
                                    Found = true;

                                    Helper.Sql_Misc_NonQuery("UPDATE communications.. Magento_Contactus SET conversationid = '" + convoid + "' WHERE ID = " + Rowid);

                                    break;
                                }
                                else
                                {
                                    if (MagetnoProductAPI.DevMode > 0)
                                    {
                                        Console.WriteLine(Preview.Contains(kustomeremail).ToString() + "; " + pageName.Contains(formType) + "; " + Preview.Replace("\n", " ").Contains(comment255));
                                        Console.WriteLine("NO");
                                    }
                                }
                            }

                            if (!Found)
                            {
                                Helper.Sql_Misc_NonQuery("UPDATE Communications..magento_contactus set Conversationid = '-1', emailsent = 100 WHERE ID = " + MagentoContactusid.ToString());
                                Helper.SendEmail("Magento_ContactUS not in Kustomer", "Magento_ContactUS not in Kustomer: ID: " + MagentoContactusid.ToString() + "; " + kustomeremail + "; TYPE: " + formType.ToString() + "; Date: " + datestamp.ToString(), "thomas@andragroup.com", "alerts@herroom.com");
                            }
                        }
                        catch (System.Exception E1)
                        {
                            Console.WriteLine("ERROR E1:" + E1.ToString());
                            bReturnvalue = false;
                        }

                    }
                    catch (System.Exception exEmail)
                    {
                        if (MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine("exEmail: " + exEmail.ToString());
                        }
                        return false;
                    }
                }
            }

            return bReturnvalue;
        }

        //2024-10-21 - NEW API
        public static string GetKustomerid(string Email)
        {
            string EndpointName;
            dynamic dJson;
            string EID = "0";
            DataTable dt;
            string Kustomerid = "";

            dt = Helper.Sql_Misc_Fetch("SELECT Isnull(KustomerID, 0) FROM communications..Kustomer_Cust WHERE Email = '" + Email + "'");
            if (dt.Rows.Count == 1 && dt.Rows[0][0].ToString() != "0")
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine("FOUND IN DB: " + dt.Rows[0][0].ToString());
                }
                return dt.Rows[0][0].ToString();
            }

            //Not in DB, go get it
            EndpointName = "https://api.kustomerapp.com/v1/customers/email=" + Email.Replace("@", "%40");
            if (MagetnoProductAPI.DevMode > 0)
            {
                Console.WriteLine(EndpointName);
            }

            var httpWebRequestEmail = (HttpWebRequest)WebRequest.Create(EndpointName);
            httpWebRequestEmail.ContentType = "application/json";
            httpWebRequestEmail.KeepAlive = true;
            httpWebRequestEmail.UserAgent = "PostmanRuntime/7.31.1";
            httpWebRequestEmail.Timeout = 60000;
            httpWebRequestEmail.Headers.Add("Authorization", "Bearer " + ConfigurationManager.AppSettings["BearerTokenKustomer"]);
            httpWebRequestEmail.Method = "GET";
            try
            {
                var httpResponseEmail = (HttpWebResponse)httpWebRequestEmail.GetResponse();
                var streamReaderEmail = new StreamReader(httpResponseEmail.GetResponseStream());
                var resultEmail = streamReaderEmail.ReadToEnd();

                dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(resultEmail);
                EID = dJson.data.id;                
                Kustomerid = EID.ToString();
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine("Custid: " + EID.ToString());
                    Console.WriteLine("Kustomerid = " + Kustomerid);
                }

                if (MagetnoProductAPI.DevMode < 2)
                {
                    if (Kustomerid != "")
                    {
                        Helper.Sql_Misc_NonQuery("EXEC Communications..[proc_mag_kustomer_cust_insert] @Email = '" + Email + "', @Kustomerid = '" + Kustomerid + "'");
                    }
                }
            }
            catch (System.Exception exEmail)
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine("ERROR: exEmail: " + exEmail.ToString());
                }
                return "ERROR: " + exEmail.ToString();
            }
        
            return Kustomerid.ToString();
        }

        //2024-10-21 - NEW API
        public static string GetConvo_Info(long MagnetoContactusId)
        {
            DataTable dt;
            string Kustomerid;
            string EndpointName;
            string comment255;
            string emailSubject;
            string SubjectSmall = "";
            dynamic dJson;
            string Preview;
            string ConvoId = "0";
            //Boolean ConvoFound = false;

            dt = Helper.Sql_Misc_Fetch("SELECT emailsubject, ISNULL(kustomeremail,'x') [kustomeremail], emailbody, ISNULL(conversationid,'') [conversationid], ordernumber, REPLACE(REPLACE(REPLACE(REPLACE(ISNULL(Comment255, 'xxx'), ' ', ''), char(10), ''), char(13), ''),'''','`') [comment255] FROM Communications..Magento_Contactus WHERE id = " + MagnetoContactusId.ToString());
            if (dt.Rows.Count == 1)
            {
                //Just Case
                Kustomerid = GetKustomerid(dt.Rows[0]["kustomeremail"].ToString());
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine("Kustomerid: " + Kustomerid + "; emailsubject: " + dt.Rows[0]["emailSubject"].ToString() + "; " + dt.Rows[0]["conversationid"].ToString());
                }

                //Get ConvoID from Kustomer if not gotten already
                if (dt.Rows[0]["conversationid"].ToString().Length > 10)
                {
                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine(dt.Rows[0]["conversationid"].ToString());
                    }
                    return dt.Rows[0]["conversationid"].ToString();
                }
                else
                {
                    comment255 = dt.Rows[0]["comment255"].ToString();
                    if (comment255.Length > 100)
                    {
                        comment255 = comment255.Substring(0, 100);
                    }
                    comment255 = comment255.Replace("'", "`");
                    comment255 = comment255.Replace("\t", "");

                    emailSubject = dt.Rows[0]["emailSubject"].ToString();
                    switch (emailSubject.ToLower())
                    {
                        case "herroom - cancel order":
                            SubjectSmall = "cancelorder";
                            break;

                        case "herroom - order status":
                            SubjectSmall = "orderstatus";
                            break;

                            //TEST ALL OF THESE
                        case "herroom - question about an item":
                            SubjectSmall = "questionaboutanitem";
                            break;

                        case "herroom - returns/exchanges":
                            SubjectSmall = "returns/exchanges";
                            break;

                        case "herroom - suggestions":
                            SubjectSmall = "suggestions";
                            break;

                        case "herRoom - technical issue / website":
                            SubjectSmall = "technicalissue/website";
                            break;

                        case "herroom contact us request":
                            SubjectSmall = "contactusrequest";
                            break;

                        case "hisroom question about":
                            SubjectSmall = "hisroomquestionabout";
                            break;

                        case "herRoom question about":
                            SubjectSmall = "herRoomquestionabout";
                            break;

                        case "your herroom order":
                            SubjectSmall = "herroomorder";
                            break;

                        case "your hisroom order":
                            SubjectSmall = "hisroomorder";
                            break;

                        //default:
                            //SubjectSmall = "";
                            //break;
                    }

                    EndpointName = "https://api.kustomerapp.com/v1/customers/" + Kustomerid + "/conversations?page=1&pageSize=100";
                    var httpWebRequestEmail = (HttpWebRequest)WebRequest.Create(EndpointName);
                    httpWebRequestEmail.ContentType = "application/json";
                    httpWebRequestEmail.KeepAlive = true;
                    httpWebRequestEmail.UserAgent = "PostmanRuntime/7.31.1";
                    httpWebRequestEmail.Timeout = 60000;
                    httpWebRequestEmail.Headers.Add("Authorization", "Bearer " + ConfigurationManager.AppSettings["BearerTokenKustomer"]);
                    httpWebRequestEmail.Method = "GET";
                    try
                    {
                        var httpResponseEmail = (HttpWebResponse)httpWebRequestEmail.GetResponse();
                        var streamReaderEmail = new StreamReader(httpResponseEmail.GetResponseStream());
                        var resultEmail = streamReaderEmail.ReadToEnd();

                        dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(resultEmail);
                        //EID = dJson.data.id;
                        //Kustomerid = EID.ToString();
                        if (MagetnoProductAPI.DevMode > 1)
                        {
                            Console.WriteLine("dJson" + dJson);
                        }

                        for (int xx = 0; xx < dJson.data.Count; xx++)
                        {
                            Preview = dJson.data[xx].attributes.preview;
                            Preview = Preview.Replace("\n", "");
                            Preview = Preview.Replace(" ", "");
                            Preview = Preview.Replace("'", "`");

                            ConvoId = dJson.data[xx].id;

                            if (MagetnoProductAPI.DevMode > 0)
                            {
                                Console.WriteLine("Preview :: " + Preview);
                            }


                            if ((Preview.ToLower().Contains("subject" + SubjectSmall) && Preview.Replace("\n", " ").Contains(comment255)) 
                                    || (Preview.ToLower().Contains("question:") && Preview.Replace("\n", " ").Contains(comment255)))
                             {
                                if (MagetnoProductAPI.DevMode > 0)
                                {
                                    Console.WriteLine("FOUND");
                                    Console.WriteLine("UPDATE communications..magento_contactus SET conversationid = '" + ConvoId + "' WHERE id = " + MagnetoContactusId);
                                }
                                if (MagetnoProductAPI.DevMode < 2)
                                {
                                    Helper.Sql_Misc_NonQuery("UPDATE communications..magento_contactus SET conversationid = '" + ConvoId + "' WHERE id = " + MagnetoContactusId);
                                }

                                return ConvoId; 
                            }
                        }
                        //NOT FOUND
                        ConvoId = "0";

                        Helper.Sql_Misc_NonQuery("UPDATE communications..magento_contactus SET conversationid = 'not found' WHERE id = " + MagnetoContactusId);


                        //if (MagetnoProductAPI.DevMode < 2)
                        //{

                        //}
                    }
                    catch (System.Exception exEmail)
                    {
                        if (MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine("ERROR: exEmail: " + exEmail.ToString());
                        }
                        return "ERROR: " + exEmail.ToString();
                    }
                }     
            }
            else
            {
                // NO ROW 
            }

            return ConvoId;
        }

        //ADD THIS TO THE API PROCESS - use MagentoContactusID for now ??
        public static Boolean PushKustomerOrderNotes(long MagentoContactId=0)
        {
            Boolean Returnvalue = true;
            DataTable dt = new DataTable();
            DataTable dtNote;
            string Convoid = "0";

            if (MagentoContactId > 0)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT kustomeremail, id, conversationid FROM communications..magento_contactus WHERE id = " + MagentoContactId.ToString());
            }
            else
            {
                // ????????????????????????
                dt = Helper.Sql_Misc_Fetch("SELECT kustomeremail, id, conversationid FROM communications..magento_contactus WHERE conversationid < 1000 AND Emailsent = 3");

            }

            foreach (DataRow dr in dt.Rows)
            {
                //1) get convid
                //2) get notes
                //3) get email
                //4)  Kustomer.AddConversationNote(Convoid, OrderNotes);

                try
                {
                    Convoid = Kustomer.GetConvo_Info(long.Parse(dr["id"].ToString()));  // this runs Kustomer.GetKustomerid()
                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("Convoid: " + Convoid);
                    }

                    if (Convoid != "0" && Convoid.Length > 10)
                    {
                        //push order info to notes on convoid 
                        dtNote = Helper.Sql_Misc_Fetch("EXEC communications..[proc_mag_customer_orders_get] '" + dr["kustomeremail"].ToString() + "'");

                        if (dtNote.Rows.Count == 1 && dtNote.Rows[0][0].ToString().Length > 1)
                        {
                           Kustomer.AddConversationNote(Convoid, "Customer Orders: " + dtNote.Rows[0][0].ToString());

                            Helper.Sql_Misc_NonQuery("UPDATE communications..magento_contactus SET conversationid = '" + Convoid + "' WHERE id = " + dr["id"].ToString() + " AND ISNULL(conversationid,0) <> '" + Convoid + "'");
                        }
                    }
                }
                catch(System.Exception ex)
                {
                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("ERROR: " + ex.ToString());
                    }
                    Returnvalue = false;
                }

            }


            return Returnvalue;
        }

    }
}
    