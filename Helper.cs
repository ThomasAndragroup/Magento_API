using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Data;
using System.Data.Odbc;
using System.Web;
using Newtonsoft.Json.Linq;
using System.Net;
using System.IO;
using Newtonsoft.Json;
using System.Text.RegularExpressions;
using System.Net.Http;
using System.Data.SqlClient;

namespace MagentoProductAPI
{
    class Helper
    {
        private static String OdbcConnectionString = ConfigurationManager.AppSettings["OdbcConnection"];
        private static String SqlConnectionString = ConfigurationManager.AppSettings["SqlConnection"];

        public static DataSet Sql_Misc_Fetch_Dataset(String Sql)
        {
            DataSet ds = new DataSet();
            SqlConnection myConnection = new SqlConnection(SqlConnectionString);
            SqlDataAdapter sDataAdapter = new SqlDataAdapter();

            try
            {
                myConnection.Open();
                SqlCommand myCommand = new SqlCommand(Sql, myConnection);
                sDataAdapter.SelectCommand = myCommand;
                sDataAdapter.Fill(ds);
                sDataAdapter.Dispose();
            }
            catch (Exception Ex)
            {
                Console.WriteLine(" ---------------- ");
                Console.WriteLine(":: " + Sql);
                Console.WriteLine(" ---------------- ");
                Console.WriteLine("ERROR: Sql_Misc_Fetch_Dataset(): " + Ex.ToString());
                Console.WriteLine(" ---------------- ");
                Console.WriteLine(Sql);
            }
            finally
            {
                myConnection.Close();
            }

            return ds;
        }


        public static DataTable Sql_Misc_Fetch(String Sql)
        {
            OdbcCommand cmd;
            DataTable dt = new DataTable();
            OdbcConnection cnn = new OdbcConnection(OdbcConnectionString);

            try
            {
                cnn.Open();
                cmd = new OdbcCommand(Sql, cnn);
                cmd.CommandTimeout = 1200;
                OdbcDataReader reader = cmd.ExecuteReader();
                dt.Load(reader);
                cmd.Dispose();
                reader.Close();
            }
            catch (Exception Ex)
            {
                Console.WriteLine(" ---------------- ");
                Console.WriteLine(":: " + Sql);
                Console.WriteLine(" ---------------- ");
                Console.WriteLine("ERROR: Sql_Misc_Fetch(): " + Ex.ToString());
                Console.WriteLine(" ---------------- ");
                Console.WriteLine(Sql);
            }
            finally
            {
                cnn.Close();
            }
            return dt;
        }

        public static Boolean Sql_Misc_NonQuery(String Sql)
        {
            OdbcCommand cmd;
            OdbcConnection cnn = new OdbcConnection(OdbcConnectionString);
            Boolean bReturnValue = true;

            //Sql = Sql.Replace("'", "''");

            if (Sql.Length > 0)
            {
                try
                {
                    cnn.Open();
                    cmd = new OdbcCommand(Sql, cnn);
                    cmd.CommandTimeout = 120000;
                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                }
                catch (Exception Ex)
                {
                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine(" ---------------- ");
                        Console.WriteLine(":: " + Sql);
                        Console.WriteLine(" ---------------- ");
                        //Console.WriteLine("ERROR: Sql_Misc_Run(): " + Ex.ToString());
                        Console.WriteLine(" ---------------- ");
                    }
                    Console.WriteLine("ERROR: Sql_Misc_Run(): " + Ex.ToString());
                    bReturnValue = false;
                }
                finally
                {
                    cnn.Close();
                }
            }

            return bReturnValue;
        }

        public static Boolean Middleware_Status_Update(long MiddlewareId, int Status)
        {
            Helper.Sql_Misc_NonQuery("UPDATE communications..middleware SET Status = " + Status.ToString() + " WHERE ID = " + MiddlewareId);

            return true;
        }

        public static Boolean Middleware_Status_Update(DataTable dt, int Status, String WhereClause = "" )
        {
            string SqlUpdate; 

            if (dt.Rows.Count == 1)
            {
                Helper.Sql_Misc_NonQuery("UPDATE communications..middleware SET Status = " + Status.ToString() + " WHERE ID = " + dt.Rows[0]["id"].ToString());
            }
            else if (dt.Rows.Count > 1)
            {
                SqlUpdate = "UPDATE communications..middleware SET Status = " + Status.ToString() + " WHERE ID IN (";
                foreach (DataRow dr in dt.Rows)
                {
                    SqlUpdate += dr["id"].ToString() + ",";
                }
                //Remove Trailing ,
                SqlUpdate = SqlUpdate.Remove(SqlUpdate.Length - 1) + ")";

                if (WhereClause.Length > 4)
                {
                    if (WhereClause.Substring(0,5).ToUpper() == "WHERE")
                    {
                        WhereClause = WhereClause.Substring(6);
                    }
                    if (WhereClause.Length > 3 && WhereClause.Substring(0,3).ToUpper() == "AND")
                    {
                        WhereClause = WhereClause.Substring(3);
                    }

                    WhereClause = WhereClause.Trim();
                    if (WhereClause.Length > 2)
                    {
                        SqlUpdate += " AND " + WhereClause;
                    }
                }
                Console.WriteLine("------- " + SqlUpdate + " --------");

                Helper.Sql_Misc_NonQuery(SqlUpdate);
            }

            return true;
        }

        public static Boolean Middleware_Status_Staging_Update(DataTable dt, int Status, String WhereClause = "")
        {
            string SqlUpdate;

            if (dt.Rows.Count == 1)
            {
                Helper.Sql_Misc_NonQuery("UPDATE communications..middleware SET LoadStaging = " + Status.ToString() + " WHERE ID = " + dt.Rows[0]["id"].ToString());
            }
            else if (dt.Rows.Count > 1)
            {
                SqlUpdate = "UPDATE communications..middleware SET LoadStaging = " + Status.ToString() + " WHERE ID IN (";
                foreach (DataRow dr in dt.Rows)
                {
                    SqlUpdate += dr["id"].ToString() + ",";
                }
                //Remove Trailing ,
                SqlUpdate = SqlUpdate.Remove(SqlUpdate.Length - 1) + ");";

                if (WhereClause.Length > 4)
                {
                    if (WhereClause.Substring(0, 5).ToUpper() == "WHERE")
                    {
                        WhereClause = WhereClause.Substring(6);
                    }
                    if (WhereClause.Length > 3 && WhereClause.Substring(0, 3).ToUpper() == "AND")
                    {
                        WhereClause = WhereClause.Substring(3);
                    }

                    WhereClause = WhereClause.Trim();
                    if (WhereClause.Length > 2)
                    {
                        SqlUpdate += " AND " + WhereClause;
                    }
                }
                Console.WriteLine("------- " + SqlUpdate + " --------");

                Helper.Sql_Misc_NonQuery(SqlUpdate);
            }

            return true;
        }


        public static string MiddlewareInsert(int Status, int WorkerId, string ToMagento, string EndpointName, string EndpointMethod, string SourceTable, string SourceId, int PreProcessorId, string Batchid = "0")
        {
            String SqlString = "";
            DataTable dt;

            try
            {
                if (Batchid.Length == 0)
                {
                    Batchid = SourceId.ToString();
                }

                SqlString = "EXEC [Communications]..[proc_mag_middleware_insert] " +
                " @Status = " + Status.ToString() +
                ", @Worker_id = " + WorkerId.ToString() +
                ", @To_Magento = '" + ToMagento + "' " +
                ", @Endpoint_name = '" + EndpointName + "' " +
                ", @Endpoint_method ='" + EndpointMethod + "' " +
                ", @Source_table ='" + SourceTable + "' " +
                ", @Source_id = '" + SourceId + "'" +
                ", @Batchid = '" + Batchid + "'" +
                ", @PreProcessorId = " + PreProcessorId.ToString();     // can be null or 0

                dt = Sql_Misc_Fetch(SqlString);
                //Batchid = dt[0]["source_id"].ToString();
                
            }
            catch (Exception Ex)
            {
                Console.WriteLine("ERROR MiddlewareInsert() " + Ex.ToString() + "\r\n" + "\r\n" + SqlString);
                Batchid = "-1";
            }

            return Batchid;
        }

        /* ----------------------------------------------------------------------------------
         * Alerts 
         * --------------------------------------------------------------------------------- */

        public static Boolean QuickEmail(String Subject = "", String Body = "", String To = "", String From = "", String CC = "", String BCC = "")
        {
            String Sql;

            // SELECT master.dbo.fn_QuickEmail (25,'herroom-com.mail.protection.outlook.com', @From, @To, @CC, @BCC, @Subject, @Body);

            Sql = "SELECT master.dbo.fn_QuickEmail(25, 'herroom-com.mail.protection.outlook.com' "
            + " , '" + From + "' "  //from
            + " , '" + To + "' "      //to
            + " , '" + CC + "' "      // CC
            + " , '" + BCC + "' "  //BCC
            + " , '" + Subject + "' "         //Subject
            + " , '" + Body + "')";        //Body 

            return Sql_Misc_NonQuery(Sql);
        }

        //PREFERRED EMAIL METHOD
        public static Boolean SendEmail( String Subject = "", String Body = "", String To = "", String From = "", String CC = "", String BCC = "")
        {
            String SqlString = "";
            Boolean ReturnValue = true;
            try
            {
                Body = Body.Replace("'", "`");

                SqlString = "EXEC [proc_mag_alert_send_error] @PreProcessorId = 0, @MiddlewareId = 0, @Process = '', @ProcessStatus = '', @ProcessEvent = ''" +
                   ", @Subject = '" + Subject + "'" +
                   ", @Body = '" + Body + "'" +
                   ", @To = '" + To + "'" +
                   ", @CC = '" + CC + "'" +
                   ", @BCC = '" + BCC + "'" +
                   ", @From = '" + From + "'";

                Sql_Misc_NonQuery(SqlString);
            }
            catch (Exception Ex)
            {
                Console.WriteLine("ERROR SendEmail() " + Ex.ToString());
                Console.WriteLine(" --------------------- ");
                Console.WriteLine("ERROR SendEmail() " + SqlString);
                ReturnValue = false;
            }

            return ReturnValue;
        }

        /// <summary>
        /// ENDPOINT_METHOD = "GET" ONLY
        /// Devmode == 10 -> mstaging 
        /// </summary>
        /// <param name="EndpointName"></param>
        /// <returns></returns>
        public static String MagentoApiGet(string EndpointName, long MiddlewareId = 0)
        {
            String Returnvalue = "";

            try
            {
                if (!EndpointName.Contains("rest"))
                {
                    if (EndpointName[0].ToString() == "/")
                    {
                        EndpointName = EndpointName.Substring(1);
                    }

                    if (MagetnoProductAPI.DevMode == 10)
                    {
                        EndpointName = "https://mstaging.herroom.com/rest/all/" + EndpointName;
                    }
                    else
                    {
                        EndpointName = "https://mcprod.herroom.com/rest/all/" + EndpointName;
                    }
                }

                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine(EndpointName + "   ----------------------------------");
                }

                var httpWebRequest = (HttpWebRequest)WebRequest.Create(EndpointName);

                httpWebRequest.ContentType = "application/json";
                httpWebRequest.Headers.Add("Authorization", "Bearer " + ConfigurationManager.AppSettings["BearerToken"]);
                httpWebRequest.KeepAlive = true;
                httpWebRequest.UserAgent = "PostmanRuntime/7.31.1";
                httpWebRequest.Timeout = 6000000;
                httpWebRequest.Method = "GET";

                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();

                //ERROR TRAPPING 
               // Console.WriteLine("RESP: " + httpResponse.StatusCode.ToString());

                if (httpResponse.StatusCode.ToString().ToString().ToLower() == "ok")
                {
                    Returnvalue = "ok";
                    if (MiddlewareId > 0)
                    {
                        Sql_Misc_NonQuery("UPDATE communications..middleware set status=600, error_message = 'OK-WAITING' WHERE id = " + MiddlewareId.ToString());
                    }

                    using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                    {
                        var result = streamReader.ReadToEnd();
                        Returnvalue = result;

                        if (MiddlewareId > 0)
                        {
                            Sql_Misc_NonQuery("UPDATE communications..middleware set status=600, from_magento = '" + result.ToString() + "' WHERE id = " + MiddlewareId.ToString());
                        }
                    }
                }
                else
                {
                    Returnvalue = "ERROR";
                }
            }
            catch(Exception ex)
            {
                Returnvalue = "ERROR: " + ex.ToString();
            }

            return Returnvalue;
        }

        public static String MagentoApiPush(String EndpointMethod, String EndpointName, String JsonBody)
        {
            string Returnvalue = "";

            try
            {
                if (EndpointName[0].ToString().Length > 0 && EndpointName[0].ToString() == "/")
                {
                    EndpointName = EndpointName.Substring(1);
                }

                EndpointName = "https://mcprod.herroom.com/rest/all/" + EndpointName;

                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine(EndpointName);
                }

                var httpWebRequest = (HttpWebRequest)WebRequest.Create(EndpointName);

                httpWebRequest.ContentType = "application/json";
                httpWebRequest.Headers.Add("Authorization", "Bearer " + ConfigurationManager.AppSettings["BearerToken"]);
                httpWebRequest.KeepAlive = true;
                httpWebRequest.UserAgent = "PostmanRuntime/7.31.1";
                httpWebRequest.Timeout = 60000;

                if (EndpointMethod.ToLower() == "get")
                {
                    httpWebRequest.Method = "GET";
                }
                else
                {
                    if (EndpointMethod.ToLower() == "put")
                    {
                        httpWebRequest.Method = "PUT";
                    }
                    else
                    {
                        httpWebRequest.Method = "POST";
                    }

                    StreamWriter sw = new StreamWriter(httpWebRequest.GetRequestStream());
                    sw.Write(JsonBody);
                    sw.Close();
                    sw.Dispose();
                }

                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();

                //ERROR TRAPPING 
                Console.WriteLine("RESP: " + httpResponse.StatusCode.ToString());

                if (httpResponse.StatusCode.ToString().ToString().ToLower() == "ok")
                {
                    using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                    {
                        var result = streamReader.ReadToEnd();
                        Returnvalue = result.ToString();        // result.Replace("'", "''")
                    }
                }
                else
                {
                    Returnvalue = "ERROR";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: " + ex.ToString());
                Returnvalue = "ERROR: " + ex.ToString();
            }

            return Returnvalue;
        }

        public static String MagentoApiPushStaging(long Middlewareid)
        {
            String Returnvalue = "";
            String json;
            DataRow dr;
            String link;
            dynamic dJson;
            string EID = "0";
            long lReturnvalue = 0;
            string Styleid;
            int stringHelper;
            string endpointName = "";

            try
            {
                DataTable dt = new DataTable();
                if (Middlewareid > 0)
                {
                    dt = Sql_Misc_Fetch("SELECT to_magento, source_table, source_id, endpoint_name, endpoint_method FROM communications..middleware with (nolock) where id = " + Middlewareid.ToString());
                }

                if (dt.Rows.Count == 1)
                {
                    json = dt.Rows[0][0].ToString();
                    dr = dt.Rows[0];
                    //Console.WriteLine("Sourceid: " + dr["source_id"].ToString());
                    //Console.WriteLine(json);

                    endpointName = dr["endpoint_name"].ToString();
                    endpointName = endpointName.Replace("hisroom/", "");
                    if (endpointName[0].ToString() == "/")
                    {
                        endpointName = endpointName.Substring(1);
                    }
                    link = "https://mstaging.herroom.com/rest/all/";


                    if (endpointName.StartsWith("all/"))
                    {
                        endpointName = endpointName.Substring(4);
                    }
                    if (endpointName.StartsWith("/all/"))
                    {
                        endpointName = endpointName.Substring(5);
                    }

                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine(link + endpointName);
                    }

                    var httpWebRequest = (HttpWebRequest)WebRequest.Create(link + endpointName);

                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.KeepAlive = true;
                    httpWebRequest.UserAgent = "PostmanRuntime/7.31.1";
                    httpWebRequest.Timeout = 600000;
                    httpWebRequest.Headers.Add("Authorization", "Bearer " + ConfigurationManager.AppSettings["BearerTokenMSTAGING"]);

                    if (dr["endpoint_method"].ToString().ToLower() == "get")
                    {
                        httpWebRequest.Method = "GET";
                    }
                    else
                    {
                        if (dr["endpoint_method"].ToString().ToLower() == "put")
                        {
                            httpWebRequest.Method = "PUT";
                        }
                        else
                        {
                            httpWebRequest.Method = "POST";
                        }

                        StreamWriter sw = new StreamWriter(httpWebRequest.GetRequestStream());
                        sw.Write(json);
                        sw.Close();
                        sw.Dispose();
                    }

                    var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();

                    //ERROR TRAPPING 
                    //Console.WriteLine("RESP: " + httpResponse.StatusCode.ToString());

                    if (httpResponse.StatusCode.ToString().ToString().ToLower() == "accepted" ||
                            httpResponse.StatusCode.ToString().ToString().ToLower() == "true")
                    {
                        //Returnvalue = "accepted";
                        Returnvalue = httpResponse.StatusCode.ToString().ToString().ToLower();
                    }
                    else if (httpResponse.StatusCode.ToString().ToString().ToLower() == "ok")
                    {
                        using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                        {
                            var result = streamReader.ReadToEnd();

                            if (result.Length > 1000)
                            {
                                Returnvalue = result.Substring(0, 1000);
                            }
                            else
                            {
                                Returnvalue = result;
                            }
                            //Console.WriteLine(Returnvalue);
                        }

                        if (dr["source_table"].ToString().ToLower() == "herroom..itemslink")
                        {
                            Sql_Misc_NonQuery("UPDATE Communications..Middleware SET LoadStaging=600 WHERE ID = " + Middlewareid.ToString());
                        }
                        else if (dr["endpoint_method"].ToString().ToLower() == "post")
                        {
                            if (long.TryParse(Returnvalue.Replace("\"", ""), out lReturnvalue))
                            {
                                EID = lReturnvalue.ToString();
                                Returnvalue = EID;
                            }
                            else
                            {
                                //Product return JSON is too complex for this...
                                if (dr["source_table"].ToString().ToLower() == "herroom..styles" || dr["source_table"].ToString().ToLower() == "herroom..items")
                                {
                                    EID = "0";
                                    if (Returnvalue.Substring(0, 100).Contains("id\":"))
                                    {
                                        stringHelper = Returnvalue.IndexOf("sku");
                                        Styleid = Returnvalue.Substring(6, stringHelper - 8);
                                        EID = Styleid;
                                        Returnvalue = Styleid;
                                    }
                                }
                                else
                                {
                                    dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(Returnvalue);

                                    if (dJson.id != null)
                                    {
                                        try
                                        {
                                            EID = dJson.id;
                                            //Console.WriteLine("EID: " + EID.ToString());
                                            Returnvalue = EID.ToString();
                                        }
                                        catch (Exception ex)
                                        {
                                            Returnvalue = "ERROR: NO MID RETURNED (1)";
                                        }
                                    }

                                    else
                                    {
                                        Returnvalue = "ERROR: NO MID RETURNED (2)";
                                    }
                                }
                            }

                            if (Middlewareid > 0)
                            {
                                if (EID == "0")
                                {
                                    Sql_Misc_NonQuery("UPDATE Communications..Middleware SET LoadStaging=700 WHERE ID = " + Middlewareid.ToString());
                                }
                                else
                                {
                                    Sql_Misc_NonQuery("UPDATE Communications..Middleware SET LoadStaging=500  WHERE ID = " + Middlewareid.ToString());
                                }
                            }
                        }
                        else if (Middlewareid > 0)
                        {
                            Sql_Misc_NonQuery("UPDATE Communications..Middleware SET LoadStaging=600 WHERE ID = " + Middlewareid.ToString());
                        }
                    }
                }
                else  //ERROR TRAPPING - need more as the error become known 
                {
                    if (Middlewareid > 0)
                    {
                        Sql_Misc_NonQuery("UPDATE Communications..Middleware SET LoadStaging=701 WHERE ID = " + Middlewareid.ToString());
                    }
                    Returnvalue = "ERROR";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: " + ex.ToString());

                if (Middlewareid > 0)
                {
                    Sql_Misc_NonQuery("UPDATE Communications..Middleware SET LoadStaging= 700 WHERE ID = " + Middlewareid.ToString());
                }

                Returnvalue = "ERROR: " + ex.ToString();
            }

            return Returnvalue;
        }


        public static String MagentoApiPush_Direct(String Method, String endpoint, String Body)
        {
            String Returnvalue = "";

            var httpWebRequest = (HttpWebRequest)WebRequest.Create(endpoint);

            httpWebRequest.ContentType = "application/json";
            httpWebRequest.KeepAlive = true;
            httpWebRequest.UserAgent = "PostmanRuntime/7.31.1";
            httpWebRequest.Timeout = 6000000;
            httpWebRequest.Headers.Add("Authorization", "Bearer " + ConfigurationManager.AppSettings["BearerToken"]);

            httpWebRequest.Method = Method;

            StreamWriter sw = new StreamWriter(httpWebRequest.GetRequestStream());
            sw.Write(Body);
            sw.Close();
            sw.Dispose();

            try
            {
                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();

                //ERROR TRAPPING 
                var streamReader = new StreamReader(httpResponse.GetResponseStream());
                var result = streamReader.ReadToEnd();

                //2024-02-13 INVENTORY UPDATES BULK
                if (httpResponse.StatusCode.ToString().ToString().ToLower() == "accepted" ||
                        httpResponse.StatusCode.ToString().ToString().ToLower() == "true" ||
                        httpResponse.StatusCode.ToString().ToString().ToLower() == "ok")
                {
                    Returnvalue = httpResponse.StatusCode.ToString().ToString().ToLower();
                }
            }
            catch (Exception exx)
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine(" -- ERROR MagentoApiPush_Direct(): " + exx.ToString());
                }

                Returnvalue = "ERROR";
            }
                    

            return Returnvalue;
        }

        public static String MagentoApiPush(long Middlewareid, string Environment = "")
           {
            String Returnvalue = "";
            String json;
            DataRow dr;
            String link;
            dynamic dJson;
            string EID = "0";
            long lReturnvalue = 0;
            string Styleid;
            int stringHelper;
            string endpointName = "";
            bool TryBoolean; 

            try
            {
                DataTable dt = new DataTable();
                if (Middlewareid > 0)
                {
                    dt = Sql_Misc_Fetch("SELECT to_magento, source_table, source_id, endpoint_name, endpoint_method FROM communications..middleware with (nolock) where id = " + Middlewareid.ToString());
                }

                if (dt.Rows.Count == 1)
                {
                    json = dt.Rows[0][0].ToString();
                    dr = dt.Rows[0];
                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("Sourceid: " + dr["source_id"].ToString());
                    }

                    endpointName = dr["endpoint_name"].ToString();
                    endpointName = endpointName.Replace("hisroom/", "");
                    if (endpointName[0].ToString() == "/")
                    {
                        endpointName = endpointName.Substring(1);
                    } 

                    if (Environment.Trim().Length > 0)
                    {
                        link = "https://" + Environment.ToLower() + ".herroom.com/rest/all/";
                    }
                    else if (dr["source_table"].ToString().ToLower() == "herroom..itemprice" || dr["source_table"].ToString().ToLower() == "Herroom..stylepricebulk")
                    {
                        link = "https://mcprod.herroom.com/rest/";
                    }
                    else
                    {
                        link = "https://mcprod.herroom.com/rest/all/";
                    }

                    if (endpointName.StartsWith("all/") )
                    {
                        endpointName = endpointName.Substring(4);
                    }
                    if ( endpointName.StartsWith("/all/"))
                    {
                        endpointName = endpointName.Substring(5);
                    }

                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine(link + endpointName);
                    }
                    var httpWebRequest = (HttpWebRequest)WebRequest.Create(link + endpointName);

                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.KeepAlive = true;
                    httpWebRequest.UserAgent = "PostmanRuntime/7.31.1";
                    httpWebRequest.Timeout = 6000000;
                    httpWebRequest.Headers.Add("Authorization", "Bearer " + ConfigurationManager.AppSettings["BearerToken"]);

                    //mstaging use
                    if (Environment.ToLower() == "mstaging")
                    {
                        string autorization = "admin" + ":" + "h_Fz5lh-lOxRicR6hQCv6CumB";
                        byte[] binaryAuthorization = System.Text.Encoding.UTF8.GetBytes(autorization);
                        autorization = Convert.ToBase64String(binaryAuthorization);
                        autorization = "Basic " + autorization;
                        httpWebRequest.Headers.Add("Authorization", autorization);
                    }

                    if (dr["endpoint_method"].ToString().ToLower() == "get")
                    {
                        httpWebRequest.Method = "GET";
                    }
                    else if (dr["endpoint_method"].ToString().ToLower() == "delete")
                    {
                        httpWebRequest.Method = "DELETE";
                    }
                    else
                    {
                        if (dr["endpoint_method"].ToString().ToLower() == "put")
                        {
                            httpWebRequest.Method = "PUT";
                        }
                        else
                        {
                            httpWebRequest.Method = "POST";
                        }

                        StreamWriter sw = new StreamWriter(httpWebRequest.GetRequestStream());
                        sw.Write(json);
                        sw.Close();
                        sw.Dispose();
                    }

                    try
                    {
                        var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();

                        //ERROR TRAPPING 
                        var streamReader = new StreamReader(httpResponse.GetResponseStream());
                        var result = streamReader.ReadToEnd();

                        //2024-02-13 INVENTORY UPDATES BULK
                        if (dr["source_table"].ToString().ToLower().Contains("herroom..sourceitems"))
                        {
                            // NO WAY TO TELL IF ERRORED ??  
                            Returnvalue = "true";
                            Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + Returnvalue + "' WHERE ID = " + Middlewareid.ToString());
                        }
                        else if (httpResponse.StatusCode.ToString().ToString().ToLower() == "accepted" ||
                                httpResponse.StatusCode.ToString().ToString().ToLower() == "true")
                        {
                            Returnvalue = httpResponse.StatusCode.ToString().ToString().ToLower();
                        }
                        else if (httpResponse.StatusCode.ToString().ToString().ToLower() == "ok")
                        {
                            //I HAVE THIS TYPE OF ERROR TRAPPING, REVISIT, MAYBE add column middleware..harvestid BOOLEAN
                            if (dr["source_table"].ToString().ToLower() == "hercust..ordercomment")
                            {
                                Returnvalue = "true";
                            }
                            else if (dr["endpoint_method"].ToString().ToLower() == "get")
                            {

                                Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + result.Replace("'", "''") + "' WHERE ID = " + Middlewareid.ToString());
                                Returnvalue = "ok";
                            }
                            else
                            {
                                Returnvalue = result;

                                if (dr["source_table"].ToString().ToLower() == "herroom..itemslink")
                                {
                                    if (result.ToString().ToUpper() == "TRUE")
                                    {
                                        Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + Returnvalue + "' WHERE ID = " + Middlewareid.ToString());
                                    }
                                    else
                                    {
                                        Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, from_magento = '" + result.ToString() + "' WHERE ID = " + Middlewareid.ToString());
                                    }
                                }
                                else if (dr["endpoint_method"].ToString().ToLower() == "post")
                                {
                                    dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(Returnvalue);

                                    if (dr["source_table"].ToString().ToLower() == "hercust..orderinvoice" || dr["source_table"].ToString().ToLower() == "hercust..ordercancel")
                                    {
                                        if (Boolean.TryParse(dJson, out TryBoolean))
                                        {
                                            Returnvalue = TryBoolean.ToString();
                                        }

                                        //if (dJson == true)
                                        //{
                                         //   Returnvalue = "true";
                                        //}
                                        //else if (dJson == false)
                                        //{
                                         //   Returnvalue = "false";
                                        //}
                                        else
                                        {
                                            Returnvalue = dJson;
                                        }
                                    }
                                    else if (dr["endpoint_name"].ToString().ToLower().Contains("all/v1/inventory/source-items"))  //2024-02-12 Inventory Bulk Update 
                                    {
                                        Returnvalue = "done, outcome unknown";
                                    }
                                    else if (long.TryParse(Returnvalue.Replace("\"", ""), out lReturnvalue))
                                    {
                                        EID = lReturnvalue.ToString();
                                        Returnvalue = EID;
                                    }

                                    //else if (dJson != null && (dJson == true || dJson == false))
                                    //{
                                    //    Returnvalue = "false";
                                    //}

                                    else if (dJson != null && dJson.id != null)
                                    {
                                        try
                                        {
                                            EID = dJson.id;
                                            Console.WriteLine("EID: " + EID.ToString());
                                            Returnvalue = EID.ToString();
                                        }
                                        catch (Exception ex)
                                        {
                                            try
                                            {
                                                Console.WriteLine("CATCH FOR EID: " + dJson);
                                            }
                                            catch
                                            {
                                                Returnvalue = "ERROR: NO MID RETURNED (0)";
                                            }
                                        }
                                    }

                                    if (long.TryParse(Returnvalue.Replace("\"", ""), out lReturnvalue))
                                    {
                                        EID = lReturnvalue.ToString();
                                        Returnvalue = EID;
                                    }
                                    else
                                    {
                                        //Product return JSON is too complex for this...
                                        if (dr["source_table"].ToString().ToLower() == "herroom..styles" || dr["source_table"].ToString().ToLower() == "herroom..items")
                                        {
                                            EID = "0";
                                            if (Returnvalue.Length > 100 && Returnvalue.Substring(0, 100).Contains("id\":"))
                                            {
                                                stringHelper = Returnvalue.IndexOf("sku");
                                                Styleid = Returnvalue.Substring(6, stringHelper - 8);
                                                EID = Styleid;
                                                Returnvalue = Styleid;
                                            }
                                            else if (Returnvalue.Contains("URL key for specified store already exists"))
                                            {
                                                Returnvalue = "URL key for specified store already exists";
                                            }
                                        }
                                        else if (dr["source_table"].ToString().ToLower() == "hercust..ordercancel")
                                        {
                                            if (Returnvalue.ToUpper() == "TRUE")
                                            {
                                                Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + Returnvalue + "' WHERE ID = " + Middlewareid.ToString());
                                            }
                                            else
                                            {
                                                Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, from_magento = '" + Returnvalue + "' WHERE ID = " + Middlewareid.ToString());

                                            }
                                        }
                                        else
                                        {
                                            dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(Returnvalue);
                                            if (dJson || !dJson)
                                            {
                                                Returnvalue = "0";
                                            }
                                            else if (dJson != null && dJson.id != null)
                                            {
                                                try
                                                {
                                                    EID = dJson.id;
                                                    Console.WriteLine("EID: " + EID.ToString());
                                                    Returnvalue = EID.ToString();
                                                }
                                                catch (Exception ex)
                                                {
                                                    Returnvalue = "ERROR: NO MID RETURNED (1)";
                                                }
                                            }
                                            else if (dr["source_table"].ToString().ToLower() == "herroom..styleoptions")
                                            {
                                                Returnvalue = "ok";
                                                EID = "100";
                                            }
                                            else
                                            {
                                                Returnvalue = "ERROR: NO MID RETURNED (2)";
                                            }
                                        }
                                    }

                                    if (Middlewareid > 0)
                                    {
                                        if (EID == "0")
                                        {
                                            Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, from_magento = '" + Returnvalue + "' WHERE ID = " + Middlewareid.ToString() + " AND Status <> 600");
                                        }
                                        else
                                        {
                                            Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=500, from_magento = '" + EID.ToString() + "' WHERE ID = " + Middlewareid.ToString());
                                        }
                                    }
                                }
                                else if (Middlewareid > 0)
                                {
                                    if (MagetnoProductAPI.DevMode > 1)
                                    {
                                        Console.WriteLine(Returnvalue);
                                    }

                                    if (Returnvalue.Length > 100)
                                    {
                                        Returnvalue = Returnvalue.Substring(0, 100);
                                    }
                                    Returnvalue = Returnvalue.Replace("'", "''");
                                    Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + Returnvalue + "' WHERE ID = " + Middlewareid.ToString() + " AND Status <> 700");
                                }
                            }
                        }
                        else  //ERROR TRAPPING - need more as the error become known 
                        {
                            Returnvalue = httpResponse.StatusCode.ToString().ToString().ToLower();
                            Console.WriteLine(Returnvalue);

                            if (Middlewareid > 0)
                            {
                                Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status= 700, error_message = 'httpResponse: " + httpResponse.StatusCode.ToString().Replace("'", "''") + "' WHERE ID = " + Middlewareid.ToString());
                            }
                            Returnvalue = "ERROR";
                        }
                    }
                    catch (WebException we)
                    {
                        Console.WriteLine(we.ToString());
                        using (var stream = we.Response.GetResponseStream())
                        using (var reader = new StreamReader(stream))
                        {
                            Returnvalue = reader.ReadToEnd();
                            Console.WriteLine(Returnvalue);                         
                        }
                        Returnvalue = Returnvalue.Replace("can't", "cannot");
                        Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=702, from_magento = '" + Returnvalue.Replace("'", "''") + "' WHERE ID = " + Middlewareid.ToString());
                    }
                }
                else
                {
                    Console.WriteLine("No Middleware/Queue Row To Process");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: " + ex.ToString());

                if (Middlewareid > 0)
                {
                    Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status= 700, error_message = 'ERROR: " + ex.ToString().Replace("'", "''") + " ' WHERE ID = " + Middlewareid.ToString());
                }

                Returnvalue = "ERROR: " + ex.ToString();
            }

            if (MagetnoProductAPI.DevMode > 1)
            {
                Console.WriteLine(Returnvalue);
                Console.WriteLine(lReturnvalue);
            }

            return Returnvalue;
        }

        //2024-06-20 IN TEST 
        public static String MagentoApiPush_Klaviyo(string APIType, String Json, string Store)
        {
            String Returnvalue = "";
            //String Json;
            //DataRow dr;
            String link;


            //    Dim HerRoomApi As New Klaviyo.Api("HXtNd6", "pk_6ef65dd2aa24cc3a5c771f1d5ea45ec500", "HGsLuE")
            //    Dim HisRoomApi As New Klaviyo.Api("MvWnUa", "pk_5823753c6de70c9928d12d639dfe9fa3d6", "HPm6jC")

            //link = "https://a.klaviyo.com/client/profiles/";
            link = ConfigurationManager.AppSettings["KLAVIYOCLIENTPROFILEHER"];  //DEFAULT FOR NOW
            if (APIType.ToUpper() == "CLIENT")
            {
                if (Store.ToUpper() == "HIS")
                {
                    link = ConfigurationManager.AppSettings["KLAVIYOCLIENTPROFILEHIS"];
                }
                else
                {
                    link = ConfigurationManager.AppSettings["KLAVIYOCLIENTPROFILEHER"];
                }
            }
            else if (APIType.ToUpper() == "EVENT" )
            {
                if (Store.ToUpper() == "HIS")
                {
                    link = ConfigurationManager.AppSettings["KLAVIYOCLIENTEVENTSHIS"];
                }
                else
                {
                    link = ConfigurationManager.AppSettings["KLAVIYOCLIENTEVENTSHER"];
                }
            }
            


            var httpWebRequest = (HttpWebRequest)WebRequest.Create(link);
            httpWebRequest.ContentType = "application/json";
            httpWebRequest.KeepAlive = true;
            httpWebRequest.UserAgent = "PostmanRuntime/7.31.1";
            httpWebRequest.Timeout = 60000;
            httpWebRequest.Headers.Add("Authorization", "pk_6ef65dd2aa24cc3a5c771f1d5ea45ec500");
            httpWebRequest.Headers.Add("revision", "2024-06-15");
            httpWebRequest.Method = "POST";

            try
            {
                StreamWriter sw = new StreamWriter(httpWebRequest.GetRequestStream());
                sw.Write(Json);
                sw.Close();
                sw.Dispose();

                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();

                //ERROR TRAPPING 
                var streamReader = new StreamReader(httpResponse.GetResponseStream());
                var result = streamReader.ReadToEnd();  // SHOULD = "1"

                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine(httpResponse.StatusCode);  //"accepted" is good
                    Console.WriteLine(Returnvalue);
                    Console.WriteLine(result);
                }
            }
            catch(Exception ex)
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine(ex.ToString());
                }

                Returnvalue = "ERROR";
            }

            return Returnvalue;
        }




        //2024-01-16
        public static String WebserviceTest()
        {
            String Returnvalue = "x";

            try
            {
                var content = new StringContent(@"{""RMANumber"":""2000001254""}", System.Text.Encoding.UTF8, "application/json");
                var url = new Uri("https://qa.herroom.com/Services.aspx/GetRMAPDFLinkforMagento");

                string posting = POSTData(content, url);           
            }
            catch (Exception exx)
            {
                Console.WriteLine(exx.ToString());
                Returnvalue = "ERROR";
            }

            return Returnvalue;
        }

        //2024-01-16
        public static string POSTData(StringContent json, Uri url)
        {
            var request = new HttpRequestMessage();
            var retResponse = new HttpResponseMessage();
            HttpClient _httpClient; // = new HttpClient();

            try
            {
                using (_httpClient = new HttpClient())
                {
                    request.Method = HttpMethod.Post;
                    request.RequestUri = url;
                    request.Content = json;

                    request.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                    var response = _httpClient.SendAsync(request).Result;
                    retResponse.StatusCode = response.StatusCode;
                    retResponse.ReasonPhrase = response.ReasonPhrase;
                    var responseData = response.Content.ReadAsStringAsync();
                    if (retResponse.StatusCode == System.Net.HttpStatusCode.OK)
                        return responseData.Result;
                    else
                        return  "Received a failure from server " + retResponse.ReasonPhrase;
                }
            }
            catch (Exception e)
            {
                 return "failed to post " + e.Message;
            }
        }

        // NOT FINISHED
        public static DataTable FetchMiddleware(int MaxRowstoReturn, long Middlewareid, int status, string Batchid, int Batchnum, string Source_Table, string EndpointName = "", string EndpointMethod = "")
        {
            DataTable dt = new DataTable();
            string SqlString;

            if (MaxRowstoReturn > 0)
            {
                SqlString = "SELECT TOP * " + MaxRowstoReturn.ToString();
            }
            else
            {
                SqlString = "SELECT * ";
            }

            if (Middlewareid > 0)
            {
                SqlString += "";
            }

            SqlString += "";
            SqlString += "";
            SqlString += "";
            SqlString += "";

            return dt;
        }

        public static Boolean Middleware_700_Retry()
        {
            Sql_Misc_NonQuery("UPDATE Middleware SET status = 100, tries = ISNULL(tries,0) + 1, posted=Dateadd(minute,10, getdate()) WHERE status in (700, 701) AND ISNULL(tries,0) < 6 AND source_table = 'Herroom..ItemsLink'");

            Sql_Misc_NonQuery("UPDATE Middleware SET status = 100, tries = ISNULL(tries,0) + 1 WHERE status in (700, 701) AND ISNULL(tries,0) < 2 AND source_table = 'hercust..ordercomment'");

            Sql_Misc_NonQuery("UPDATE Middleware SET status = 100, tries = ISNULL(tries,0) + 1 WHERE status in (700, 701) AND ISNULL(tries,0) < 6 AND source_table NOT IN ('hercust..ordercomment','Herroom..ItemsLink') ");

            return true;
        }

        public static string HtmlToPlainText(string html)
        {
            const string tagWhiteSpace = @"(>|$)(\W|\n|\r)+<";//matches one or more (white space or line breaks) between '>' and '<'
            const string stripFormatting = @"<[^>]*(>|$)";//match any character between '<' and '>', even when end tag is missing
            const string lineBreak = @"<(br|BR)\s{0,1}\/{0,1}>";//matches: <br>,<br/>,<br />,<BR>,<BR/>,<BR />
            var lineBreakRegex = new Regex(lineBreak, RegexOptions.Multiline);
            var stripFormattingRegex = new Regex(stripFormatting, RegexOptions.Multiline);
            var tagWhiteSpaceRegex = new Regex(tagWhiteSpace, RegexOptions.Multiline);

            var text = html;
            //Decode html specific characters
            text = System.Net.WebUtility.HtmlDecode(text);
            //Remove tag whitespace/line breaks
            text = tagWhiteSpaceRegex.Replace(text, "><");
            //Replace <br /> with line breaks
            text = lineBreakRegex.Replace(text, Environment.NewLine);
            //Strip formatting
            text = stripFormattingRegex.Replace(text, string.Empty);

            return text;
        }

    }
}

