using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using Newtonsoft.Json;
using Magento_MCP;
using System.Net.Http;
using System.Configuration;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace MagentoProductAPI
{
    class Orders
    {
        public static Boolean Process_RMA_Incoming(long MiddlewareIncomingid = 0)
        {
            Boolean bReturnvalue = true;
            System.Data.DataTable dt = new System.Data.DataTable();
            Magento_MCP.MagentoModels.SalesModels.RmaIncoming RMA;
            List<Magento_MCP.MagentoModels.HelperModels.RmaItem> RMAItems;


            if (MiddlewareIncomingid > 0)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT id, from_magento FROM Communications..middleware_incoming WHERE endpoint_method = 'INCOMING - Module RMA' AND id = " + MiddlewareIncomingid.ToString());
            }
            else
            {
                dt = Helper.Sql_Misc_Fetch("SELECT id, from_magento FROM Communications..middleware_incoming WHERE Posted = 0 AND endpoint_method = 'INCOMING - Module RMA'; ");
            }

            Helper.Middleware_Status_Update(dt, 200);

            foreach (DataRow dr in dt.Rows)
            {
                RMA = new Magento_MCP.MagentoModels.SalesModels.RmaIncoming();

                RMA = JsonConvert.DeserializeObject<Magento_MCP.MagentoModels.SalesModels.RmaIncoming>(dr["from_magento"].ToString(), new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
               ///2024-06-24 broke RMAItems = RMA.RmaDataObject.Items;

            }

            return bReturnvalue;
        }


        public static Boolean Process_Order_Invoice(long Middlewareid = 0)
        {
            Boolean bReturnvalue = true;
            System.Data.DataTable dt = new System.Data.DataTable();
            String APIReturn = "";
            long mid;

            try
            {
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE source_table = 'hercust..orderinvoice' AND id = " + Middlewareid.ToString());
                }
                else
                {
                    Helper.Sql_Misc_NonQuery("UPDATE communications..middleware SET status = 100 FROM Middleware WHERE status = 200 AND source_table = 'hercust..orderinvoice' AND posted < dateadd(hour, 3, getdate())");

                    dt = Helper.Sql_Misc_Fetch("SELECT TOP 1000 id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND source_table = 'hercust..orderinvoice' ORDER BY ID ");
                }

                Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                {
                    APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");

                    if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                    {
                        //Middleware should already by set to 500
                        Console.WriteLine("ERROR Process_Order_Invoice: " + dr["source_id"].ToString());
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, from_magento = 'NO INVOICEID... " + APIReturn + "' WHERE ID = " + dr["id"].ToString());

                        Order_AddHTNote("Middleware", "API Invoice Failed, no InvoiceId was received from Magento; please invoice manually.", long.Parse(dr["source_id"].ToString()), false, false);
                    }
                    else //Should return new value
                    {
                        if (long.TryParse(APIReturn, out mid))
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + mid.ToString() + "' WHERE ID = " + dr["id"].ToString());

                            //do THIS FOR ALL and update sprocs to use BatchID and set batchnum = null
                            Helper.Sql_Misc_NonQuery("UPDATE communications..Orderids SET mag_invoiceid = " + mid.ToString() + ", LoadStepNumber = 3 WHERE (OrderIncrementId = '" + dr["source_id"].ToString() + "' OR Orderno = " + dr["source_id"].ToString() + ")");

                            Helper.Sql_Misc_NonQuery("UPDATE Hercust..Orders SET CCIntegrityCheck = '" + mid.ToString() + "' WHERE (MagentoOrderno = '" + dr["source_id"].ToString() + "' OR Orderno = " + dr["source_id"].ToString() + ") AND ISNULL(CCIntegrityCheck,'') = '' ");
                        }
                        else
                        {
                            if (APIReturn.ToLower() == "braintree_googlepay")
                            {
                                Helper.Sql_Misc_NonQuery("UPDATE communications..Orderids SET mag_invoiceid = 1000, LoadStepNumber = 3 WHERE (OrderIncrementId = '" + dr["source_id"].ToString() + "' OR Orderno = " + dr["source_id"].ToString() + ")");

                                Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + APIReturn + "' WHERE ID = " + dr["id"].ToString());
                            }
                            else
                            {
                                Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=701, from_magento = 'NO INVOICEID... " + APIReturn.Replace("can't", "cannot") + "' WHERE ID = " + dr["id"].ToString());
                                Order_AddHTNote("Middleware", "API Invoice Failed, no InvoiceId was received from Magento; please invoice manually.", long.Parse(dr["source_id"].ToString()), false, false);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Process_Order_Invoice(" + Middlewareid.ToString() + ": " + ex.ToString());
                bReturnvalue = false;
            }
            return bReturnvalue;
        }


        public static Boolean Process_Order_Ship(long Middlewareid = 0)
        {
            Boolean bReturnvalue = true;
            System.Data.DataTable dt = new System.Data.DataTable();
            String APIReturn = "";
            long mid;

            try
            {
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE source_table = 'hercust..ordership' AND id = " + Middlewareid.ToString());
                }
                else
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT TOP 200 id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND source_table = 'hercust..ordership' ORDER BY ID");
                }

                Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                {
                    APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");

                    if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                    {
                        //Middleware should already by set to 500
                        Console.WriteLine("ERROR Process_Order_Invoice: " + dr["source_id"].ToString());
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, from_magento = 'NO SHIPID... " + APIReturn + "' WHERE ID = " + dr["id"].ToString());

                        Order_AddHTNote("Middleware", "API Invoice Failed, no Shipping Id was received from Magento.", long.Parse(dr["source_id"].ToString()), false, false);
                    }
                    else //Should return new value
                    {
                        if (long.TryParse(APIReturn, out mid))
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + mid.ToString() + "' WHERE ID = " + dr["id"].ToString());

                            Helper.Sql_Misc_NonQuery("UPDATE communications..Orderids SET mag_shipid = " + mid.ToString() + ", LoadStepNumber = 5 WHERE Orderno = " + dr["source_id"].ToString());

                            //if all is shipped, do this:
                            Process_Order_Ship_Received(dr["source_id"].ToString());
                         }
                        else
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, from_magento = 'NO SHIPID... " + APIReturn + "' WHERE ID = " + dr["id"].ToString());
                            Order_AddHTNote("Middleware", "API Invoice Failed, no Shipping Id was received from Magento.", long.Parse(dr["source_id"].ToString()), false, false);

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Process_Order_Ship(" + Middlewareid.ToString() + ": " + ex.ToString());
                bReturnvalue = false;
            }
            return bReturnvalue;
        }

        public static Boolean Process_Order_Ship_Received(string Orderno)
        {
            Boolean bReturnvalue = true;
            System.Data.DataTable dt;
            System.Data.DataTable dtInc;

            dt = Helper.Sql_Misc_Fetch("SELECT COUNT(*) [zero_count] FROM Orderids WHERE OrderIncrementId IN (SELECT OrderIncrementId FROM Orderids WHERE Orderno = " + Orderno + ") AND ISNULL(Mag_shipid, 0) = 0; ");
            dtInc = Helper.Sql_Misc_Fetch("SELECT Magentoid[OrderIncrementId] FROM HERCUST..ORDERS WHERE Orderno = " + Orderno + ";");

            if (dtInc.Rows.Count == 1)
            {
                if (dt.Rows.Count == 1 && dt.Rows[0]["zero_count"].ToString() == "0")
                {
                    //2028-08-29 'complete' = 'Order Shipped'
                    //Orders.Order_AddComment("Magento shipping id(s) processed in HerTools.", long.Parse(dtInc.Rows[0]["OrderIncrementId"].ToString()), "middleware_complete", 0, long.Parse(Orderno));
                    Orders.Order_AddComment("Magento shipping id(s) processed in HerTools.", long.Parse(dtInc.Rows[0]["OrderIncrementId"].ToString()), "complete", 0, long.Parse(Orderno));
                }
            }
            return bReturnvalue;
        }

        /// TTRIBBLE NEW: 2023-11-27 FROM MCP Code Order_Cancel()
        /// Hercust..OrderNotes Trigger creates middleware row status 100, do NOT process until status 101
        /// <param name="Middlewareid"></param>
        /// <returns></returns>
        public static Boolean Prep_Order_Cancel(long Middlewareid = 0)
        {
            Boolean bReturnValue = true;
            System.Data.DataTable dtMiddlewareRows;
            System.Data.DataTable dt;
            DataRow dr;

            if (Middlewareid > 0)
            {
                dtMiddlewareRows = Helper.Sql_Misc_Fetch("SELECT id, source_id FROM Communications..Middleware WHERE source_table = 'hercust..ordercancel' AND id = " + Middlewareid.ToString());
            }
            else
            {
                dtMiddlewareRows = Helper.Sql_Misc_Fetch("SELECT id, source_id FROM Communications..Middleware WHERE status = 100 AND source_table = 'hercust..ordercancel' AND posted <= Getdate();");
            }

            try
                { 
                Helper.Middleware_Status_Update(dtMiddlewareRows, 200);

                foreach (DataRow drMiddlewareRow in dtMiddlewareRows.Rows)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT OO.magentoid, OO.OrderNo, OO.Cancel, ISNULL(PO.OrderNo,0) [po], ISNULL(PO.BackOrder,0) [po_bo], ISNULL(PO.Shipped,0) [po_shipped] "
                    + ", ISNULL(CO.orderno,0) [co], ISNULL(CO.BackOrder,0) [co_bo], ISNULL(CO.Shipped,0) [co_shipped], ISNULL(CO.Cancel,0) [co_cancel] FROM hercust..orders OO "
                    + " LEFT OUTER JOIN hercust..orders PO ON PO.OrderNo = OO.ParentOrder "
                    + " LEFT OUTER JOIN hercust..orders CO ON CO.OrderNo = OO.ChildOrder WHERE OO.orderno = " + drMiddlewareRow["source_id"].ToString());

                    if (dt.Rows.Count > 0)
                    {
                        dr = dt.Rows[0];

                        if (dr["cancel"].ToString() == "1" || dr["cancel"].ToString().ToLower() == "true")   //NOT SHIPPED
                        {
                            if (int.Parse(dr["co"].ToString()) == 0 || dr["co"].ToString().ToLower() == "false" || dr["co_cancel"].ToString() == "1" || dr["co_cancel"].ToString().ToLower() == "true")
                            {
                                Helper.Sql_Misc_NonQuery("UPDATE communications..middleware SET status = 101 WHERE id = " + drMiddlewareRow["id"].ToString());
                            }
                            else if (dr["co_shipped"].ToString() == "1" || dr["co_shipped"].ToString().ToLower() == "true")
                            {
                                Helper.Sql_Misc_NonQuery("UPDATE communications..middleware SET status = 101 WHERE id = " + drMiddlewareRow["id"].ToString());
                            }
                        }
                        else if (int.Parse(dr["po"].ToString()) > 0 || int.Parse(dr["co"].ToString()) > 0)
                        {
                            if ((dr["po_shipped"].ToString() == "0" && dr["co_shipped"].ToString() == "0") || (dr["po_shipped"].ToString().ToLower() == "false" && dr["co_shipped"].ToString().ToLower() == "false"))
                            {
                                //UPDATE MIddlewareRows, leave status alone...set Posted = Getdate() + 4 hours, tries++   
                                Helper.Sql_Misc_NonQuery("UPDATE communications..middleware SET Posted=DateAdd(hour, 4 ,Getdate()) WHERE id = " + drMiddlewareRow["id"].ToString());
                            }
                            else    //Push Order Cancel to Magento
                            {
                                Helper.Sql_Misc_NonQuery("UPDATE communications..middleware SET status = 101 WHERE id = " + drMiddlewareRow["id"].ToString());
                            }
                        }
                        else  // NO CHILD ORDER
                        {
                            if (dr["po_shipped"].ToString() == "0" || dr["po_shipped"].ToString().ToLower() == "false")
                            {
                                Helper.Sql_Misc_NonQuery("UPDATE communications..middleware SET status = 101 WHERE id = " + drMiddlewareRow["id"].ToString());
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("ERROR: Order_Cancel(): N0=o Orders Record Found for Order/ID: " + drMiddlewareRow["source_id"].ToString());
                        Helper.Sql_Misc_NonQuery("UPDATE communications..middleware SET status = 700, error_message = 'No Orders Record Found for Order' WHERE id = " + drMiddlewareRow["id"].ToString());
                        bReturnValue = false;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Order_Cancel(): " + ex.ToString());
                Helper.Middleware_Status_Update(dtMiddlewareRows, 100, "Status=200 AND Status <> 700"); 

                bReturnValue = false;
            }
            return bReturnValue;
        }


        /// Hercust..OrderNotes Trigger creates middleware row status 100, do NOT process until status 101
        public static Boolean Process_Order_Cancel(long Middlewareid = 0)
        {
            System.Data.DataTable dt;
            String APIReturn = "";
            long mid;

            // FOR NOW, DO NOTHING
            if (Middlewareid > 0)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE source_table = 'hercust..ordercancel' AND id = " + Middlewareid.ToString());
            }
            else
            {
                dt = Helper.Sql_Misc_Fetch("SELECT TOP 40 id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND source_table = 'hercust..ordercancel' ORDER BY ID");
            }

            Helper.Middleware_Status_Update(dt, 200);

            foreach (DataRow dr in dt.Rows)
            {
                APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");

                if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                {
                    //Middleware should already by set to 500
                    Console.WriteLine("ERROR Process_Order_Cancel: " + dr["source_id"].ToString());
                    Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, from_magento = 'CANCELLATION ERROR... " + APIReturn.Replace("'", "`") + "' WHERE ID = " + dr["id"].ToString());
                }
                else //Should return new value
                {
                    if (long.TryParse(APIReturn, out mid))
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + mid.ToString() + "' WHERE ID = " + dr["id"].ToString());
                    }
                    else
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, from_magento = 'NO SHIPID... " + APIReturn + "' WHERE ID = " + dr["id"].ToString());
                    }
                }
            }

            return true;
        }

        public static Boolean Process_Order_Refund(long Middlewareid = 0)
        {
            System.Data.DataTable dt;
            String APIReturn = "";
            long mid;

            if (Middlewareid > 0)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] " +
                    ", (SELECT ISNULL(Magentoid, '0') FROM Hercust..Orders with(nolock) " +
                    "       WHERE Orderno = Convert(varchar(32), ISNULL(MM.batchid, '0')))[MagentoId] " +
                    " FROM Communications..Middleware MM with(nolock) " +
                    " WHERE Posted <= Getdate() AND Status = 100 AND source_table = 'hercust..orderrefund' AND ID = " + Middlewareid.ToString());
            }
            else
            {
                dt = Helper.Sql_Misc_Fetch("SELECT TOP 200 id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] " +
                    ", (SELECT ISNULL(Magentoid, '0') FROM Hercust..Orders with(nolock) " +
                    "       WHERE Orderno = Convert(varchar(32), ISNULL(MM.batchid, '0'))) [MagentoId] " +
                    " FROM Communications..Middleware MM with(nolock) " +
                    " WHERE Posted <= Getdate() AND Status = 100 AND source_table = 'hercust..orderrefund' ORDER BY ID");
            }

            Helper.Middleware_Status_Update(dt, 200);

            foreach (DataRow dr in dt.Rows)
            {
                APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");

                if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                {
                    //Middleware should already by set to 500
                    Console.WriteLine("ERROR Process_Order_Refund: " + dr["source_id"].ToString());
                    Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, from_magento = 'REFUND ERROR... " + APIReturn + "' WHERE ID = " + dr["id"].ToString());

                    Order_AddHTNote("Middleware", "API Refund/Invoice Failed, no refund was received from Magento; please refund manually.", long.Parse(dr["batchid"].ToString()), false, false);
                    Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700 WHERE ID = " + dr["id"].ToString());

                }
                else //Should return new value
                {
                    if (long.TryParse(APIReturn, out mid))
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + mid.ToString() + "' WHERE ID = " + dr["id"].ToString());

                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Orderids SET Mag_creditmemoid  = " + mid + " WHERE Mag_invoiceid = " + dr["source_id"].ToString());

                        Helper.Sql_Misc_NonQuery("UPDATE communications..Magento_Refunds SET Processed=1, State = 2 " +
                            " FROM communications..Magento_Refunds MR " +
                            " INNER JOIN Communications..Orderids OI ON OI.orderno = MR.Orderno WHERE OI.Mag_invoiceid = " + dr["source_id"].ToString());

                        Order_AddHTNote("Middleware", "Magento Refund processed by System/CSR", long.Parse(dr["batchid"].ToString()), false, false);
                    }
                    else
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=800, from_magento = 'NO SHIPID... " + APIReturn + "' WHERE ID = " + dr["id"].ToString());

                        Order_AddHTNote("Middleware", "Magentod Refund did NOT process, an error occurred during refund process.", long.Parse(dr["batchid"].ToString()), false, false);
                    }
                }
            }

            return true;
        }

        public static Boolean Process_Order_Comments(long Middlewareid = 0)
        {
            Boolean bReturnvalue = true;
            System.Data.DataTable dt;
            String APIReturn = "";
            //long mid;

            if (Middlewareid > 0)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,'') [batchid]  " +
                    " FROM Communications..Middleware MM with(nolock) " +
                    " WHERE Posted <= Getdate() AND Status = 100 AND source_table = 'hercust..ordercomment' AND ID = " + Middlewareid.ToString());
            }
            else
            {
                dt = Helper.Sql_Misc_Fetch("SELECT TOP 20 id, source_id, endpoint_method, ISNULL(batchid,'') [batchid]  " +
                    " FROM Communications..Middleware MM with(nolock) " +
                    " WHERE Posted <= Getdate() AND Status = 100 AND source_table = 'hercust..ordercomment' ORDER BY 1");
            }

            Helper.Middleware_Status_Update(dt, 200);

            foreach (DataRow dr in dt.Rows)
            {
                APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");

                if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                {
                    //Middleware should already by set to 500
                    Console.WriteLine("ERROR Process_Order_Comments: " + dr["source_id"].ToString());
                    Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, from_magento = 'API ERROR... " + APIReturn + "' WHERE ID = " + dr["id"].ToString());
                }
                else //Should return new value
                {
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + APIReturn + "' WHERE ID = " + dr["id"].ToString());
                    }
                }

                System.Threading.Thread.Sleep(1500);
            }

            return bReturnvalue;
        }

        //DO NOT USE YET
        public static long Process_Order_Missing_Report()
        {
            System.Data.DataTable dt;
            String APIReturn = "";

            Helper.Sql_Misc_NonQuery("EXEC communications..[proc_mag_report_orders_notinmiddlware_insert_get];");

            dt = Helper.Sql_Misc_Fetch("SELECT TOP 1 id FROM Communications..Middleware with(nolock) WHERE source_table = 'get orders' AND status IN (100) ORDER BY id desc");
            if (dt.Rows.Count == 1)
            {
                Console.WriteLine("MID: " + dt.Rows[0][0].ToString());

                APIReturn = Helper.MagentoApiPush(long.Parse(dt.Rows[0][0].ToString()));
                Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status = 600, from_magento = '" + APIReturn.Replace("'", "''") + "' WHERE ID = " + dt.Rows[0][0].ToString());

                return long.Parse(dt.Rows[0][0].ToString());
            }
            else
            {
                return -1;
            }
            //return 0;
        }

        public static Boolean Order_AddComment(string Comment, long MagentoOrderid, string Status, int isPublic = 1, long OrderNo = 0)
        {
            //NEED SPROC TO CREATE, INSERT JSON INTO MIDDLEWARE OR CALL hercust..ordercomment()
            string Json;

            Comment = Comment.Replace("'", "''");
            Comment = Comment.Replace("\"", "''");

            Json = @"{""statusHistory"": { " +
                    @"""comment"": """ + Comment + @""", " +
                    @"""created_at"": null, " +
                    @"""entity_id"": 0," +
                    @"""entity_name"": null," +
                    @"""is_customer_notified"": 0," +
                    @"""is_visible_on_front"": " + isPublic.ToString() + ", " +
                    @"""parent_id"": " + MagentoOrderid + "," +
                    @"""status"": """ + Status + @"""," +    
                    @"""extension_attributes"": null }}";

            Console.WriteLine(Json);

            if (OrderNo > 0)
            {
                Helper.MiddlewareInsert(100, 20, Json, "V1/orders/" + MagentoOrderid + "/comments", "POST", "hercust..ordercomment", OrderNo.ToString(), 0);
            }
            else
            {
                Helper.MiddlewareInsert(100, 20, Json, "V1/orders/" + MagentoOrderid + "/comments", "POST", "hercust..ordercomment", MagentoOrderid.ToString(), 0);
            }

            return true;
        }

        public static Boolean Order_AddHTNote(string Author, string Message, long HTOrderNo, Boolean isPublic, bool isAction)
        {
            int iPublic = 0;
            int iAction = 0;

            if (isPublic)
            {
                iPublic = 1;
            }
            if (isAction)
            {
                iAction = 1;
            }

            Helper.Sql_Misc_NonQuery("INSERT Hercust..OrderNotes(OrderNo, Posted, Author, Message, IsPublic, IsAction) SELECT " + HTOrderNo.ToString() + ", GETDATE(), '" + Author + "', '" + Message + "', " + iPublic.ToString() + ", " + iAction.ToString());

            return true;
        }

        public static Boolean Missing_Order_Report_API(long Middlewareid = 0, Boolean RunReport = false)
        {
            Boolean bReturnvalue = true;
            System.Data.DataTable dt;
            String APIReturn = "";

            try
            {
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT Top 1 id FROM communications..Middleware WHERE source_table = 'get orders' AND ID = " + Middlewareid.ToString());
                }
                else
                {
                     dt = Helper.Sql_Misc_Fetch("EXEC communications..[proc_mag_report_orders_notinmiddlware_insert_get]");
                }

                Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                {
                    Console.WriteLine(dr[0].ToString());
                    //Helper.MagentoApiPush() handles all of this for GETs
                    APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");

                    if (RunReport)
                    {
                        Helper.Sql_Misc_NonQuery("EXEC communications..[proc_mag_report_orders_notinmiddlware] @MidToUse = " + dr["id"].ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Missing_Order_Report_API(" + Middlewareid.ToString() + "): " + ex.ToString());
                bReturnvalue = false;
            }

            return bReturnvalue;
        }


        public static Boolean Missing_Order_Report(long Middlewareid = 0, Boolean SendEmail=false)
        {
            System.Data.DataTable dt;

            if (SendEmail)
            {
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("EXEC Communications..[proc_mag_report_orders_notinmiddlware] @MidToUse = " + Middlewareid.ToString() + ", @Debugging = 0");
                }
                else
                {
                    dt = Helper.Sql_Misc_Fetch("DECLARE @mid bigint; " +
                    " SELECT @mid = Max(id) FROM Communications..Middleware " +
                    " WHERE source_table = 'get orders' AND status IN (500, 600) ORDER BY 1 desc; " +
                    " EXEC communications..[proc_mag_report_orders_notinmiddlware] @MidToUse = @MID, @DeBugging = 0;");
                }
            }
            else
            {
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("EXEC Communications..[proc_mag_report_orders_notinmiddlware] @MidToUse = " + Middlewareid.ToString() + ", @Debugging = 1");
                }
                else
                {
                    dt = Helper.Sql_Misc_Fetch("DECLARE @mid bigint; " +
                    " SELECT @mid = Max(id) FROM Communications..Middleware " +
                    " WHERE source_table = 'get orders' AND status IN(500, 600) ORDER BY 1 desc; " +
                    " EXEC communications..[proc_mag_report_orders_notinmiddlware] @MidToUse = @MID, @DeBugging = 1;");
                }

                //what to do with this?
                if (dt.Rows.Count == 1)
                {
                    Console.WriteLine("Missing Orders: " + dt.Rows[0][1].ToString());
                }
            }

            return true;
        }

        public static Boolean Order_Comments_DB()
        {
            System.Data.DataTable dt;
            dt = Helper.Sql_Misc_Fetch("SELECT * FROM communications..Magento_order_comments WHERE Status = 100 AND Posted < getdate()");
            foreach(DataRow dr in dt.Rows)
            {
                Console.WriteLine("SourceID: " + dr["source_id"].ToString());
                Helper.Sql_Misc_NonQuery("UPDATE communications..Magento_order_comments SET status = 600 WHERE id = " + dr["id"].ToString()); 
            }

            return true;
        }

        public static Boolean PayPal_Invoice_Order_Update(long Middlewareid = 0)
        {
            Boolean bReturnvalue = true;
            System.Data.DataTable dt;

            // 1) insert Get row to Middleware table 
            //      EXEC communications..[proc_mag_orderinvoice_paypal_htorder_getinsert] :: returns middlewareid
            // 2) process get row
            //      Helper.MagentoApiPush(xxx);
            // 3) Process new sproc supplying middleware row id
            //      EXEC communications..[proc_mag_orderinvoice_paypal_htorder_update] @Middlewareid =  xx

            try
            {
                if (Middlewareid == 0)
                {
                    dt = Helper.Sql_Misc_Fetch("EXEC communications..[proc_mag_orderinvoice_paypal_htorder_getinsert]");
                    if (dt.Rows.Count == 1)
                    {
                        Middlewareid = long.Parse(dt.Rows[0]["Middewareid"].ToString());
                    }
                }

                if (Middlewareid > 0)
                {
                    if (Helper.MagentoApiPush(Middlewareid).ToLower() != "error")
                    {
                        Helper.Sql_Misc_Fetch("EXEC communications..[proc_mag_orderinvoice_paypal_htorder_update] @Middlewareid = " + Middlewareid.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("PayPal_Invoice_Order_Update() ERROR::" + ex.ToString());
                bReturnvalue = false;
            }

            return bReturnvalue;
        }

        public static Boolean Invoice_Order_Get_Update(long Middlewareid = 0)
        {
            Boolean bReturnvalue = true;
            System.Data.DataTable dt;
            System.Data.DataTable dtOrder;
            dynamic dJson;
            dynamic dJsonItem;
            dynamic dJsonItemAll;
            string magItemids;
            System.Data.DataTable dtItems;
            System.Data.DataTable dtItemsP;
            int magQty;
            int magCount;
            Boolean PriceMismatch;

            if (Middlewareid > 0)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT id, from_magento, status FROM communications..middleware with (nolock) WHERE source_table = 'get invoices' AND id = " + Middlewareid.ToString());
            }
            else
            {
                dt = Helper.Sql_Misc_Fetch("SELECT TOP 1 id, from_magento, status FROM communications..middleware with (nolock) WHERE source_table = 'get invoices' AND status = 600; ");
            }

            try
            {
                if (dt.Rows.Count > 0)  //use first row only
                {
                    dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(dt.Rows[0]["from_magento"].ToString());

                    if (MagetnoProductAPI.DevMode > -1)
                    {
                        Console.WriteLine("COUNT: " + dJson.items.Count.ToString());
                    }
    
                    for (int xx = 0; xx < dJson.items.Count; xx++)
                    {
                        if (MagetnoProductAPI.DevMode > -1 && xx%100 == 0)
                        {    
                            Console.WriteLine(" -- " + xx.ToString());
                        }

                        dJsonItem = dJson.items[xx];
                        if (MagetnoProductAPI.DevMode > 1)
                        {
                            Console.WriteLine(xx.ToString() + " : " + dJsonItem.increment_id);
                        }

                        dtOrder = Helper.Sql_Misc_Fetch("SELECT OO.orderno, ISNULL(CCIntegrityCheck,'') [CCIntegrityCheck], shipped, round(Isnull(OO.total,0),2) [total], ROUND(Isnull(OO.total,0),2) + 0.30 [total2], ROUND(Isnull(OO.total,0),2) - 0.30 [total3] ,  ROUND((ship*taxrate) + total,2) [totalshiptax] ,  ROUND((ship*taxrate),2) + total  + 0.50 [totalshiptax2], ROUND((ship * taxrate), 2) + total - 0.50 [totalshiptax3] "
                            + " FROM hercust..orders OO with (nolock) "
                            + " WHERE oo.magentoid = " + dJsonItem.order_id + " AND HELD=0 AND CANCEL=0 AND SHIPPED=0 AND ISNULL(CCIntegrityCheck,'') = '' AND (CCAuth='PAYPAL' OR BACKORDER=0)  ");

                    //    +" WHERE oo.magentoid = " + dJsonItem.order_id + " AND BACKORDER=0 AND HELD=0 AND CANCEL=0 AND SHIPPED=0 AND ISNULL(CCIntegrityCheck,'') = '' ");


                    foreach (DataRow drOrder in dtOrder.Rows)
                        {
                            try
                            {
                                if (drOrder["CCIntegrityCheck"].ToString() == dJsonItem.entity_id.ToString())
                                {
                                    // DONE, NOTHING TO DO
                                    if (MagetnoProductAPI.DevMode > 1)
                                    {
                                        Console.WriteLine("Order Has CCIntegrityCheck already: " + drOrder["orderno"].ToString() + " = " + dJsonItem.entity_id);
                                    }
                                }
                                else if (drOrder["CCIntegrityCheck"].ToString() == "" && (drOrder["total"].ToString() == dJsonItem.grand_total.ToString() || (Decimal.Parse(dJsonItem.grand_total.ToString()) > Decimal.Parse(drOrder["total3"].ToString()) && Decimal.Parse(dJsonItem.grand_total.ToString()) < Decimal.Parse(drOrder["total2"].ToString()))))
                                {
                                    // Update Orders..CCIntegrityCheck
                                    if (MagetnoProductAPI.DevMode > 0)
                                    {
                                        Console.WriteLine("UPDATE Hercust..Orders SET CCIntegrityCheck = '" + dJsonItem.entity_id + "' WHERE Orderno = " + drOrder["orderno"].ToString());
                                    }
                                    else
                                    {
                                        Console.WriteLine("UPDATED: Hercust..Orders CCIntegrityCheck = '" + dJsonItem.entity_id + "' WHERE Orderno = " + drOrder["orderno"].ToString());
                                        Helper.Sql_Misc_NonQuery("UPDATE Hercust..Orders SET CCIntegrityCheck = '" + dJsonItem.entity_id + "' WHERE Orderno = " + drOrder["orderno"].ToString() + " AND CANCEL=0");
                                    }
                                }
                                else if (drOrder["CCIntegrityCheck"].ToString() == "" && (drOrder["totalshiptax"].ToString() == dJsonItem.grand_total.ToString() || (Decimal.Parse(dJsonItem.grand_total.ToString()) > Decimal.Parse(drOrder["totalshiptax3"].ToString()) && Decimal.Parse(dJsonItem.grand_total.ToString()) < Decimal.Parse(drOrder["totalshiptax2"].ToString()))))
                                {

                                    if (MagetnoProductAPI.DevMode > 0)
                                    {
                                        Console.WriteLine("UPDATE Hercust..Orders SET CCIntegrityCheck = '" + dJsonItem.entity_id + "' WHERE Orderno = " + drOrder["orderno"].ToString());
                                    }
                                    else
                                    {
                                        Console.WriteLine("UPDATED: Hercust..Orders CCIntegrityCheck = '" + dJsonItem.entity_id + "' WHERE Orderno = " + drOrder["orderno"].ToString());
                                        Helper.Sql_Misc_NonQuery("UPDATE Hercust..Orders SET CCIntegrityCheck = '" + dJsonItem.entity_id + "' WHERE Orderno = " + drOrder["orderno"].ToString() + " AND CANCEL=0");
                                    }
                                }
                                else if (drOrder["CCIntegrityCheck"].ToString() == "")
                                {
                                    // GO THROUGH ITEMS on Items and json, if all magentoids are present, use invoice id ??
                                    //dtItems = Helper.Sql_Misc_Fetch("");

                                    // !! BEGIN 2024-05-06 need to check itemids, if all in for a order, okay it
                                    magItemids = "";
                                    magQty = 0;
                                    magCount = 0;
                                    PriceMismatch = true;
                                    dtItems = new System.Data.DataTable();  // NEEDED FOR LOOP 
                                    dtItemsP = new System.Data.DataTable();  // NEEDED FOR LOOP 
                                    dJsonItemAll = dJson.items[xx].items;

                                    for (int yy = 0; yy < dJsonItemAll.Count; yy++)
                                    {
                                        if (MagetnoProductAPI.DevMode > 0)
                                        {
                                            Console.WriteLine(dJsonItemAll[yy] + " : " + dJsonItemAll[yy].entity_id + "; orderno: " + drOrder["orderno"].ToString());
                                        }
                                        //magItemids += dJsonItemAll[yy].entity_id + ",";
                                        magItemids += dJsonItemAll[yy].order_item_id + ",";
                                        magQty += int.Parse(dJsonItemAll[yy].qty.ToString());
                                        magCount++;
                                    }

                                    if (magItemids != "")
                                    {
                                        magItemids = magItemids.Substring(0, magItemids.Length - 1);
                                        //Console.WriteLine(magItemids + "; magqty: " + magQty.ToString());
                                        //Console.WriteLine("SELECT COUNT(*) FROM Hercust..Items II INNER JOIN HERCUST..Orders OO ON ordernum=orderno AND ISNULL(ccintegritycheck, '') = '' WHERE Ordernum = " + drOrder["orderno"].ToString() + " AND (Magentoid IN (" + magItemids + ") OR Parentid in (" + magItemids + "));");
                                        //Console.WriteLine("SELECT COUNT(*) [cnt], (SUM(Qty) * 2) [qty] FROM Hercust..Items WHERE Ordernum = " + drOrder["orderno"].ToString() + " AND (Magentoid IN (" + magItemids + ") OR ParentItemid in (" + magItemids + "));");

                                        dtItems = Helper.Sql_Misc_Fetch("SELECT COUNT(*) * 2 [cnt], ISNULL((SUM(Qty) * 2),0) [qty] FROM Hercust..Items WHERE Ordernum = " + drOrder["orderno"].ToString() + " AND (Magentoid IN (" + magItemids + ") OR ParentItemid in (" + magItemids + "));");
                                        dtItemsP = Helper.Sql_Misc_Fetch("SELECT COUNT(*) * 2 [cnt], ISNULL((SUM(Qty) * 2),0) [qty] FROM Hercust..Items II INNER JOIN Hercust..Orders OO ON Orderno = Ordernum AND OO.Magentoid = " + dJsonItem.order_id + " AND (II.Magentoid IN (" + magItemids + ") OR II.ParentItemid in (" + magItemids + "));");

                                    }

                                    if (MagetnoProductAPI.DevMode > 0)
                                    {
                                        Console.WriteLine(":::: " + dtItems.Rows[0]["cnt"].ToString() + " == " + dJsonItemAll.Count + " && " + dtItems.Rows[0]["qty"].ToString() + " == " + magQty.ToString());
                                        Console.WriteLine(":::: " + dtItemsP.Rows[0]["cnt"].ToString() + " == " + dJsonItemAll.Count + " && " + dtItemsP.Rows[0]["qty"].ToString() + " == " + magQty.ToString());
                                    }

                                    if (dtItems.Rows.Count > 0)
                                    {
                                        Console.WriteLine("-- " + dtItems.Rows[0]["cnt"].ToString() + "; QTY: " + dtItems.Rows[0]["qty"].ToString() + " ; dJsonItemAll.Count: " + magCount.ToString() + " Magcnt:" + magQty.ToString());
                                        Console.WriteLine("-- ITEMP COUNT: " + dtItemsP.Rows[0]["cnt"].ToString() + "; QTY: " + dtItemsP.Rows[0]["qty"].ToString() + " ; dtItemsP.Rows.Count: " + dtItemsP.Rows.Count.ToString() + " Magcnt:" + magQty.ToString());

                                        //if (int.Parse(dtItems.Rows[0]["cnt"].ToString()) == int.Parse(dtItems.Rows.Count.ToString()) && int.Parse(dtItems.Rows[0]["qty"].ToString()) == magQty)
                                        if (int.Parse(dtItems.Rows[0]["cnt"].ToString()) == int.Parse(dJsonItemAll.Count.ToString()) && int.Parse(dtItems.Rows[0]["qty"].ToString()) == magQty)
                                        {
                                            PriceMismatch = false;
                                            Console.WriteLine(" ___ INVOICE IS GOOD ___ UPDATED: Hercust..Orders CCIntegrityCheck = '" + dJsonItem.entity_id + "' WHERE Orderno = " + drOrder["orderno"].ToString());
                                            Helper.Sql_Misc_NonQuery("UPDATE Hercust..Orders SET CCIntegrityCheck = '" + dJsonItem.entity_id + "' WHERE Orderno = " + drOrder["orderno"].ToString() + " AND CANCEL = 0 AND ISNULL(CCIntegrityCheck,'') = '' ");
                                            Helper.Sql_Misc_NonQuery("INSERT Hercust..OrderNotes(OrderNo, Posted, Author, Message, IsPublic, IsAction) SELECT " + drOrder["orderno"].ToString() + ", GETDATE(), 'API Process', 'Invoice id captured from Magento', 0, 0");
                                        }
                                        else if (dtItemsP.Rows.Count > 0 && int.Parse(dtItems.Rows[0]["cnt"].ToString()) > 0)
                                        {
                                            if (int.Parse(dtItemsP.Rows[0]["cnt"].ToString()) == int.Parse(dJsonItemAll.Count.ToString()) && int.Parse(dtItemsP.Rows[0]["qty"].ToString()) == magQty)
                                            {
                                                PriceMismatch = false;
                                                Console.WriteLine(" ___ INVOICE IS GOOD 2 ___ UPDATED: Hercust..Orders CCIntegrityCheck = '" + dJsonItem.entity_id + "' WHERE Orderno = " + drOrder["orderno"].ToString());
                                                Helper.Sql_Misc_NonQuery("UPDATE Hercust..Orders SET CCIntegrityCheck = '" + dJsonItem.entity_id + "' WHERE Orderno = " + drOrder["orderno"].ToString() + " AND CANCEL = 0 AND ISNULL(CCIntegrityCheck,'') = '' ");
                                                Helper.Sql_Misc_NonQuery("INSERT Hercust..OrderNotes(OrderNo, Posted, Author, Message, IsPublic, IsAction) SELECT " + drOrder["orderno"].ToString() + ", GETDATE(), 'API Process', 'Invoice id captured from Magento', 0, 0");
                                            }
                                        }
                                    }

                                    //2024-05-06 END

                                    if (MagetnoProductAPI.DevMode > -1 && PriceMismatch)
                                    {
                                        Console.WriteLine("Price Mismatch '" + drOrder["orderno"].ToString() + " :: " + drOrder["total"].ToString() + " <> " + dJsonItem.grand_total.ToString());
                                    }

                                }
                            }
                            catch(Exception exloop)
                            {
                                if (MagetnoProductAPI.DevMode > 0)
                                {
                                    Console.WriteLine("--ERROR Invoice_Order_Get_Update() INNER LOOP: " + exloop.ToString());
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine("--ERROR Invoice_Order_Get_Update(): " + ex.ToString());
                }

                bReturnvalue = false;
            }

            return bReturnvalue;
        }

        //2024-08-07
        public static Boolean AbandondedCartsReport_Process()
        {
            Boolean bReturnvalue = true;
            DateTime nn = DateTime.Now;
            System.Data.DataTable dt;
            string datetouse;

            //nn = nn.AddDays(-1);
            datetouse = nn.Year.ToString();
            if (nn.Month < 10)
            {
                datetouse += "-0" + nn.Month.ToString();
            }
            else
            {
                datetouse += "-" + nn.Month.ToString();
            }
            if (nn.Day < 10)
            {
                datetouse += "-0" + nn.Day.ToString();
            }
            else
            {
                datetouse += "-" + nn.Day.ToString(); 
            }

            //Console.WriteLine("SELECT getdate(), 101, 101, 'GET', 'https://mcprod.herroom.com/rest/all/V1/carts/search?searchCriteria[filter_groups][0][filters][0][field]=created_at&searchCriteria[filter_groups][0][filters][0][value]=" + datetouse + "T12:00:01Z&searchCriteria[filter_groups][0][filters][0][condition_type]=gt'");

            // INSERT communications..middleware(posted, status, worker_id, endpoint_method, endpoint_name)
            // SELECT getdate(), 101, 101, 'GET', 'https://mcprod.herroom.com/rest/all/V1/carts/search?searchCriteria[filter_groups][0][filters][0][field]=created_at&searchCriteria[filter_groups][0][filters][0][value]=2024-08-07T12:00:01Z&searchCriteria[filter_groups][0][filters][0][condition_type]=gt'

            Helper.Sql_Misc_NonQuery("INSERT communications..middleware(posted, status, worker_id, endpoint_method, endpoint_name, source_table) SELECT getdate(), 101, 101, 'GET', 'all/V1/carts/search?searchCriteria[filter_groups][0][filters][0][field]=created_at&searchCriteria[filter_groups][0][filters][0][value]=" + datetouse + "T00:00:01Z&searchCriteria[filter_groups][0][filters][0][condition_type]=gt', 'abandondedcarts'");

            dt = Helper.Sql_Misc_Fetch("SELECT Max(ID) [id] FROM communications..middleware WHERE endpoint_method = 'get' and source_table = 'abandondedcarts'");
            if (dt.Rows.Count > 0)
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine("Middelware GET ID: " + dt.Rows[0]["id"].ToString());
                }
                Helper.MagentoApiPush(int.Parse(dt.Rows[0]["id"].ToString()));
                //Insert Get into middleware
                //Process api row
                //AbandondedCartsReport_Fetch(id)
                AbandondedCartsReport_Fetch(int.Parse(dt.Rows[0]["id"].ToString()));

                /// AbandondedCartsReport_Send();
            }

            return bReturnvalue;
        }

        public static Boolean AbandondedCartsReport_Fetch(int Middlewareid)
        {
            Boolean bReturnvalue = true;
            System.Data.DataTable dt;
            dynamic dJson;
            dynamic dJsonItems;
            dynamic dJsonSubItems;
            int zz = 0;
            decimal ItemTotal;
            string CustomerEmail;
            long CustomerId;

            dt = Helper.Sql_Misc_Fetch("SELECT * FROM Communications..Middleware WHERE id = " + Middlewareid.ToString());

            if (dt.Rows.Count > 0)
            {
                //Helper.Sql_Misc_NonQuery("Truncate Table Communications..AbandondedCarts");

                dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(dt.Rows[0]["from_magento"].ToString());
                dJsonItems = dJson.items;

                for (int xx = 0; xx < dJsonItems.Count; xx++)
                {
                    ItemTotal = 0;
                    //ItemCount = 0;

                    // reserved_order_id
                    if (dJsonItems[xx].reserved_order_id == null)
                    {
                        //Console.WriteLine("NO ID: " + dJsonItems[xx].id + "; " + dJsonItems[xx].items.Count);
                        
                        if (dJsonItems[xx].items_count > 0 && dJsonItems[xx].items != null)
                        {
                            for (int yy = 0; yy < dJsonItems[xx].items.Count; yy++)
                            {
                                //    ItemTotal += decimal.Parse(dJsonItems[xx].items[yy].price.ToString());

                                if (dJsonItems[xx].customer.email != null)
                                {
                                    if (MagetnoProductAPI.DevMode > 0)
                                    {
                                        Console.WriteLine("email " + dJsonItems[xx].customer.email + "; " + dJsonItems[xx].customer.id);
                                    }
                                    CustomerEmail = dJsonItems[xx].customer.email;
                                    CustomerId = dJsonItems[xx].customer.id;
                                }
                                else
                                {
                                    CustomerEmail = "";
                                    CustomerId = 0;
                                }

                                Helper.Sql_Misc_NonQuery("EXEC Communications..[proc_mag_abandondedcarts_insert] @sku = '" + dJsonItems[xx].items[yy].sku + "', @Qty = " + dJsonItems[xx].items[yy].qty.ToString() + ", @Skuname = '" + dJsonItems[xx].items[yy].name.ToString().Replace("'", "`") + "', @Cartid = '" + dJsonItems[xx].id + "'"
                                         + ", @CustomerEmail = '" + CustomerEmail + "', @Customerid = " + CustomerId.ToString());

                            }

                            //Console.WriteLine("NO ID: " + dJsonItems[xx].items_count + "; QTY: " + dJsonItems[xx].items_qty + "; total: " + ItemTotal.ToString());

                           
                        }
                    }
                    //else
                    //{
                    //    Console.WriteLine("ID " + dJsonItems[xx].reserved_order_id);
                    //}

                    zz = 1;
                }
            }


            return bReturnvalue;
        }

        //2024-09-24
        //Get total for order, json, if totals are off, report it
        public static Boolean OrderSumCheck(string MagentoOrderno)
        {
            Boolean ReturnValue = true;
            System.Data.DataTable dt;
            System.Data.DataTable dtItems;
            dynamic Json;
            string Fetch;
            decimal oTotal = 0;
            decimal oMerchtot = 0;
            decimal oTaxtot = 0;
            decimal oSalestaxamt = 0;

            List<decimal> jItemQtys = new List<decimal>();
            List<string> jSkus = new List<string>();
            int ItemCount = 0;
            int SkuMatchindex;
            string Magentoid;
            string ErrorFound = "";

            // dt = Magento_MCP.HTFunctions.Sql_Misc_Fetch("SELECT [Orderno], [total], [merchtot], [taxtot], [saletaxamt] FROM Hercust..Orders WHERE MagenotOrderno = '" + MagentoOrderno + "'");
            // foreach (DataRow dr in dt.Rows)
            //{
            // }

            dt = Helper.Sql_Misc_Fetch("SELECT max(magentoid) [magentoid], round(sum(total),2) [total], sum(merchtot) [merchtot], sum(taxtot) [taxtot], sum(salestaxamt) [salestaxamt] FROM Hercust..Orders WHERE MagentoOrderno = '" + MagentoOrderno + "'");
            if (dt.Rows.Count == 1)
            {
                //dtItems = Helper.Sql_Misc_Fetch("SELECT II.ordernum, II.qty, II.total, II.Magentoid, II.ParentItemId FROM hercust..Orders OO with(nolock) INNER JOIN Hercust..ITEMS II with(nolock) ON Orderno = Ordernum WHERE OO.MagentoOrderno = '" + MagentoOrderno + "'");
                //dtItems = Helper.Sql_Misc_Fetch("SELECT SUM(II.qty) [qty], Sum(II.total) [Total], II.Magentoid [Magentoid], [ParentItemId] FROM hercust..Orders OO with(nolock) INNER JOIN Hercust..ITEMS II with(nolock) ON Orderno = Ordernum WHERE OO.MagentoOrderno = '" + MagentoOrderno + "' GROUP BY Ordernum, II.Magentoid, [ParentItemId]  ");
                dtItems = Helper.Sql_Misc_Fetch("SELECT SUM(II.qty) [qty], Round(Sum(II.total),2) [Total], max(II.Magentoid) [Magentoid], max([ParentItemId]) FROM hercust..Orders OO with(nolock) INNER JOIN Hercust..ITEMS II with(nolock) ON Orderno = Ordernum WHERE OO.MagentoOrderno = '" + MagentoOrderno + "' GROUP BY II.Magentoid, [ParentItemId]  ");

                oTotal = decimal.Parse(dt.Rows[0]["total"].ToString());
                oMerchtot = decimal.Parse(dt.Rows[0]["merchtot"].ToString());
                oTaxtot = decimal.Parse(dt.Rows[0]["taxtot"].ToString());
                oSalestaxamt = decimal.Parse(dt.Rows[0]["salestaxamt"].ToString());
                Magentoid = dt.Rows[0]["magentoid"].ToString();

                Fetch = Helper.MagentoApiGet("V1/orders/?searchCriteria[filter_groups][0][filters][0][field]=entity_id&searchCriteria[filter_groups][0][filters][0][value]=" + Magentoid + "&searchCriteria[filter_groups][0][filters][0][condition_type]=eq", 0);
                Json = Newtonsoft.Json.JsonConvert.DeserializeObject(Fetch);

                if (Json.items.Count > 0)
                {
                    //Console.WriteLine(Json.items[0].items.Count.ToString());
                    for (int jj = 0; jj < Json.items[0].items.Count; jj++)
                    {
                        if (Json.items[0].items[jj].product_type == "simple")
                        {
                            jItemQtys.Add(decimal.Parse(Json.items[0].items[jj].qty_ordered.ToString()));
                            jSkus.Add(Json.items[0].items[jj].item_id.ToString());
                            ItemCount++;
                        }
                    }

                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine(decimal.Round(oTotal, 2).ToString() + "; " + oMerchtot.ToString() + "; " + oTaxtot.ToString() + "; " + oSalestaxamt.ToString());

                        Console.WriteLine(Json.items.Count.ToString());
                        Console.WriteLine("Json total: " + Json.items[0].base_grand_total + "; ");

                        for (int ss = 0; ss < ItemCount; ss++)
                        {
                            Console.WriteLine(jSkus[ss] + "; " + jItemQtys[ss].ToString());
                        } 
                    }

                    if (decimal.Round(decimal.Parse(Json.items[0].base_grand_total.ToString()),0) != decimal.Round(decimal.Parse(dt.Rows[0]["total"].ToString()),0))
                    {
                        ErrorFound = MagentoOrderno + " TOTAL ISSUE: HT: " + dt.Rows[0]["total"].ToString() + " <> " + Json.items[0].base_grand_total.ToString() + " ^^";
                        Console.WriteLine(" ---- "  + " : "  + MagentoOrderno + " TOTAL ISSUE: HT: " + dt.Rows[0]["total"].ToString() + " <> " + Json.items[0].base_grand_total.ToString());
                    }

                    foreach (System.Data.DataRow drItem in dtItems.Rows)
                    {             
                        //SkuMatch = jSkus.Find(drItem["Magentoid"].ToString().Equals);
                        SkuMatchindex = jSkus.FindIndex(drItem["Magentoid"].ToString().Equals);

                        if (SkuMatchindex >= 0)
                        {
                            if (MagetnoProductAPI.DevMode > 1)
                            {
                                Console.WriteLine("SKU " + drItem["Magentoid"].ToString() + " QTY HT: " + drItem["Qty"].ToString() + "; " + jItemQtys[SkuMatchindex].ToString());
                            }

                            if (int.Parse(drItem["Qty"].ToString()) != int.Parse(jItemQtys[SkuMatchindex].ToString()))
                            {
                                Console.WriteLine("ISSUE: " + MagentoOrderno + " - SKU " + drItem["Magentoid"].ToString() + " QTY HT: " + drItem["Qty"].ToString() + "; " + jItemQtys[SkuMatchindex].ToString());


                                ErrorFound += "ISSUE: " + MagentoOrderno + " - SKU " + drItem["Magentoid"].ToString() + " QTY HT: " + drItem["Qty"].ToString() + "; " + jItemQtys[SkuMatchindex].ToString() + " ^^";
                                //Console.WriteLine("ISSUE: " + drItem["ordernum"].ToString() + " - SKU " + drItem["Magentoid"].ToString() + " QTY HT: " + drItem["Qty"].ToString() + "; " + jItemQtys[SkuMatchindex].ToString());
                            }
                        }
                    }

                    if (ErrorFound.Length > 0)
                    {
                        //Console.WriteLine(ErrorFound);
                        Console.WriteLine("-------------------------------------------------");
                    }

                }
                else
                {
                    Console.WriteLine("Json ERROR !!!!");
                }
            }

            return ReturnValue;
        }

        public static Boolean OrderInvoiceFetchProcessing(int MaxRowstoProcess = 150)
        {
            Boolean ReturnValue = true;
            System.Data.DataTable dtO;

            dtO = Helper.Sql_Misc_Fetch("SELECT TOP " + MaxRowstoProcess.ToString() + " CCIntegrityCheck, ccauth, MagentoOrderNo, Magentoid "
                + " FROM hercust..orders OO with(nolock) WHERE cancel = 0 "
                + " AND shipped = 0 AND siteversion = 1 and shipped = 0 and isnull(CCIntegrityCheck, '') = '' "
                + " AND ISNULL(captureattempts, 0) = 0 ORDER BY orderno");

            foreach (DataRow drO in dtO.Rows)
            {
                Console.WriteLine(drO["MagentoOrderNo"].ToString() + ", " + drO["Magentoid"].ToString());
                Orders.OrderInvoiceFetch(drO["MagentoOrderNo"].ToString(), drO["Magentoid"].ToString());
            }

            return ReturnValue;
        }

        public static Boolean OrderInvoiceFetch(string MagentoOrderno, string MagentoId)
        {
            Boolean ReturnValue = true;
            //System.Data.DataTable dtId;
            //System.Data.DataTable dtData = new System.Data.DataTable();
            string APIString;
            dynamic dJson;
            string CCId = "-1";

            try
            {
                APIString = Helper.MagentoApiGet("V1/invoices?searchCriteria[filter_groups][0][filters][0][field]=order_id&searchCriteria[filter_groups][0][filters][0][value]=" + MagentoId + "&searchCriteria[filter_groups][0][filters][0][condition_type]=eq");
                dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(APIString);

                if (dJson.items.Count == 0 || dJson.items[0].entity_id == null)
                {
                    Console.WriteLine("NULL");
                }
                else
                {
                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        //Console.WriteLine(dJson.ToString());
                        Console.WriteLine("INVID: " + dJson.items[0].entity_id);
                    }

                    CCId = dJson.items[0].entity_id.ToString();
                }

                if (CCId.Length > 4)
                {
                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("UPDATE Hercust..Orders SET CCintegrityCheck = '" + CCId + "' WHERE Magentoorderno = '" + MagentoOrderno + "' AND ISNULL(CCintegritycheck,'') = '' ;");

                        Console.WriteLine("UPDATE Communications..Orderids SET mag_invoiceid = 1002 WHERE OrderIncrementId = '" + MagentoOrderno + "' AND ISNULL(OrderIncrementId,0) = 0;");
                    }

                    Helper.Sql_Misc_NonQuery("UPDATE Hercust..Orders SET CCintegrityCheck = '" + CCId + "' WHERE Magentoorderno = '" + MagentoOrderno + "' AND ISNULL(CCintegritycheck,'') = '' ;");
                    Helper.Sql_Misc_NonQuery("UPDATE Communications..Orderids SET mag_invoiceid = 1002, LoadStepNumber = 3 WHERE OrderIncrementId = '" + MagentoOrderno + "' AND ISNULL(mag_invoiceid,0) = 0 ;");
                }
                else
                {
                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("UPDATE Hercust..Orders SET CaptureAttempts = (ISNULL(CaptureAttempts,0) + 1) WHERE Magentoorderno = '" + MagentoOrderno + "' AND ISNULL(CCintegritycheck,'') = '' ;");
                    }

                    Helper.Sql_Misc_NonQuery("UPDATE Hercust..Orders SET CaptureAttempts = (ISNULL(CaptureAttempts,0) + 1) WHERE Magentoorderno = '" + MagentoOrderno + "' AND ISNULL(CCintegritycheck,'') = '' ;");
                }
            }
            catch (Exception ex)
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine("-- ERROR OrderInvoiceFetch(" + MagentoOrderno + "): " + ex.ToString());
                }
                ReturnValue = false;
            }

            return ReturnValue;
        }

        //2024-10-29 WORKS
        public static Boolean OrderRefundAPI(int Orderno, int Entityid)
        {
            Boolean Returnvalue = true;
            string APIReturn = "";
            dynamic dJson;
            dynamic dJsonMainItem;
            dynamic dJsonItems;
            System.Data.DataTable dt;

            if (Entityid == 0 && Orderno > 0)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT magentoid FROM Hercust..Orders WHERE Orderno = " + Orderno.ToString());

                if (dt.Rows.Count == 1)
                {
                    Entityid = int.Parse( dt.Rows[0]["magentoid"].ToString());
                }
            }

            try
            {
                APIReturn = Helper.MagentoApiGet("V1/orders/?searchCriteria[filter_groups][0][filters][0][field]=entity_id&searchCriteria[filter_groups][0][filters][0][value]=" + Entityid + "&searchCriteria[filter_groups][0][filters][0][condition_type]=eq");
                if (MagetnoProductAPI.DevMode > 1)
                {
                    Console.WriteLine(APIReturn);
                }

                if (APIReturn.Length > 100)
                {
                    dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(APIReturn);
                    dJsonMainItem = dJson.items;
                    if (dJsonMainItem.Count == 1)
                    {
                        dJsonItems = dJsonMainItem[0].items;
                    }
                    else
                    {
                        return false;
                    }

                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("dJsonMainItem = " + dJsonMainItem.Count.ToString());
                        Console.WriteLine("dJsonItems = " + dJsonItems.Count.ToString());
                    }

                    //Show Order as updated despite finding REfunds
                    Helper.Sql_Misc_NonQuery("UPDATE Communications..Magento_Refunds_API SET DateUpdated = Getdate() WHERE orderno IN (SELECT orderno FROM hercust..orders WHERE magentoid = " + Entityid  + ")");

                    for (int yy = 0; yy < dJsonItems.Count; yy++)
                    {
                        if (dJsonItems[yy].product_type == "simple")
                        {
                            if (dJsonItems[yy].qty_refunded > 0 && MagetnoProductAPI.DevMode < 2)
                            {
                                if (MagetnoProductAPI.DevMode > 0)
                                {
                                    Console.WriteLine("SKU: " + dJsonItems[yy].sku + "; " + dJsonItems[yy].qty_refunded + "; " + dJsonItems[yy].parent_item.amount_refunded + "; " + dJsonItems[yy].item_id + "; TOTAL: " + dJsonMainItem[0].base_total_refunded);
                                }

                                Helper.Sql_Misc_NonQuery("EXEC Communications..[proc_mag_magentorefundsapi_update] " 
                                + " @Entityid = " + Entityid.ToString()
                                + " , @itemid = " + dJsonItems[yy].item_id
                                + " , @QtyRefunded = " + dJsonItems[yy].qty_refunded
                                + " , @AmtRefunded = " + dJsonItems[yy].parent_item.amount_refunded
                                + " , @TotalAmtRefunded = " + dJsonMainItem[0].base_total_refunded);
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine("ERROR :: " + ex.ToString());
                }

                Returnvalue = false;
            }

            return Returnvalue;
        }

    }

    //public static Boolean AbandondedCartsReport_Send()
    //{
    //    Boolean bReturnvalue = true;

        // need sproc to compile and email this

    //    return bReturnvalue;
    //}


    //rmaPost 
    class RMA
    {
        class rmaPost
        {
            [JsonProperty("rmaPost")]
            public RMAShippingLabel RMAShippingLabel { get; set; }
        }

        private class RMAShippingLabel
        {
            [JsonProperty("increment_id")]
            public string increment_id { get; set; }

            [JsonProperty("status_id")]
            public string status_id { get; set; }

            [JsonProperty("pdf_attachment")]
            public string pdf_attachment { get; set; }

        }

        //2024-09-27
        //Dateto use is like "2024-09-01" format (yyyy-mm-dd)
        public static Boolean Fetch_RMA_MiddlewareIncomingInsert(string DateBegin = "", string DateEnd = "")
        {
            Boolean Returnvalue = true;
            dynamic dJson;
            string APIReturn;
            string Day;
            string Search;
           
            try
            {
                if (DateBegin != "" && DateEnd != "")
                {
                    Search = "V1/mst_rma/search?searchCriteria[filter_groups][0][filters][0][field]=created_at&searchCriteria[filter_groups][0][filters][0][value]=" + DateBegin + "T00:00&searchCriteria[filter_groups][0][filters][0][condition_type]=gt"
                    + "&searchCriteria[filter_groups][1][filters][0][field]=created_at&searchCriteria[filter_groups][1][filters][0][value]=" + DateEnd + "T23:59&searchCriteria[filter_groups][1][filters][0][condition_type]=lt";
                }
                else
                {
                    Day = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Day.ToString();
                    Search = "V1/mst_rma/search?searchCriteria[filter_groups][0][filters][0][field]=created_at&searchCriteria[filter_groups][0][filters][0][value]=" + Day + "T00:00&searchCriteria[filter_groups][0][filters][0][condition_type]=gt";
                }

                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine(Search);
                }

                //APIReturn = Helper.MagentoApiGet("V1/mst_rma/search?searchCriteria[filter_groups][0][filters][0][field]=created_at&searchCriteria[filter_groups][0][filters][0][value]=" + Day + "T00:00&searchCriteria[filter_groups][0][filters][0][condition_type]=gt");
                APIReturn = Helper.MagentoApiGet(Search);
                if (APIReturn.Length > 50)
                {
                    dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(APIReturn);
                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine(dJson.Count.ToString());
                    }
                    for (int aa = 0; aa < dJson.Count; aa++)
                    {
                        // JUST INSERT EACH INTO Middleware_Incoming
                        Helper.Sql_Misc_NonQuery("INSERT Communications..Middleware_Incoming(posted, status, worker_id, from_magento, endpoint_name, endpoint_method, source_table, source_id) "
                            + " SELECT getdate(), 100, 3333, '" + dJson[aa].ToString().Replace("'", "''") + "', 'INCOMING', 'RMA FETCH', 'RMA', '" + dJson[aa].order_increment_id + ";" + dJson[aa].rma.increment_id + "' "
                            + " WHERE 0 = (SELECT COUNT(*) FROM Communications..Middleware_Incoming WHERE source_table = 'RMA' AND source_id = '" + dJson[aa].order_increment_id + ";" + dJson[aa].rma.increment_id + "')");
                    }
                }
            }
            catch (Exception ex)
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine("ERROR: Fetch_RMA_MiddlewareIncomingInsert(): " + ex.ToString());
                }
                Returnvalue = false;
            }

            return Returnvalue;
        }

        public static Boolean Process_RMA_MiddlewareIncomingTable(int MaxRowstoProcess = 20)
        {
            Boolean Returnvalue = true;
            System.Data.DataTable dt;

            dt = Helper.Sql_Misc_Fetch("SELECT TOP " + MaxRowstoProcess.ToString() + " id, from_magento, source_id FROM Communications..Middleware_Incoming WHERE source_table = 'RMA' AND status = 100 ORDER BY id");
            foreach(System.Data.DataRow dr in dt.Rows)
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine("To Process, Row: " + dr["id"].ToString() + " = " + dr["source_id"].ToString());
                }

                Process_Incoming_DirectRMA_Insert_3(long.Parse(dr["id"].ToString()));
              }

            return Returnvalue;
        }

        //2024-09-27
        public static Boolean Process_Incoming_DirectRMA_Insert_3(long Middlewareid = 0, int MiddlewareIncomingStatus = 100)
        {
            Boolean Returnvalue = true;
            System.Data.DataTable dt;
            System.Data.DataTable dtOrdernums;
            System.Data.DataTable dtRMAs;
            System.Data.DataTable dtOrderItems;
            System.Data.DataTable dtOrderItemIds;
            String OrderItemReturnsIds;
            String sHelper;
            List<long> Ordernums = new List<long>();
            String SqlString1;
            dynamic dJson;
            string OrderNotesMessage = "";

            //load all info to arrays 
            if (Middlewareid > 0)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT id, from_magento FROM Communications..middleware_incoming WHERE id = " + Middlewareid.ToString() + " AND endpoint_method = 'RMA FETCH' ");
            }
            else
            {
                dt = Helper.Sql_Misc_Fetch("SELECT id, from_magento FROM Communications..middleware_incoming WHERE status = " + MiddlewareIncomingStatus.ToString() + " AND endpoint_method = 'RMA FETCH' ");
            }

            foreach (DataRow dr in dt.Rows)
            {
                sHelper = "";
                try
                {
                    if (MagetnoProductAPI.DevMode == 0)
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE communications..middleware_incoming SET status = 200, worker_id = 7777 WHERE ID = " + dr["id"].ToString());
                    }

                    dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(dr["from_magento"].ToString());
                    //dJsonItems = dJson.rma.items;

                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("CNT: " + dJson.items.Count.ToString());
                    }

                    //Pass 1 for item, see how many orderno are involved for split orders
                    for (int zz = 0; zz < dJson.items.Count; zz++)
                    {
                        sHelper += dJson.items[zz].order_item_id.ToString() + ',';

                        //if (zz < dJson.items[zz].Count - 1)
                        //{
                        //    sHelper += ",";
                        //}
                    }

                    if (sHelper.EndsWith(","))
                    {
                        sHelper = sHelper.Substring(0, sHelper.Length - 1);
                    }

                    OrderNotesMessage = dJson.return_label_type;

                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("Items COUNT: " + dJson.items.Count.ToString());
                        Console.WriteLine("sHelper = " + sHelper + ", dJson.rma.created_at:" + dJson.rma.created_at);
                        Console.WriteLine("SELECT DISTINCT Ordernum, 0 [RmaId] FROM herCust..items WHERE ParentItemId IN(" + sHelper + ") OR Magentoid IN(" + sHelper + ") ORDER BY 1");
                    }

                    //dtOrdernums = Helper.Sql_Misc_Fetch("SELECT DISTINCT Ordernum, 0 [RmaId], '' [ONMessage] FROM herCust..items WHERE ParentItemId IN (" + sHelper + ") OR Magentoid IN (" + sHelper + ") ORDER BY 1");
                    dtOrdernums = Helper.Sql_Misc_Fetch("SELECT DISTINCT Ordernum, 0 [RmaId] FROM herCust..items WHERE ParentItemId IN (" + sHelper + ") OR Magentoid IN (" + sHelper + ") ORDER BY 1");

                    dtOrdernums.Columns["RmaId"].ReadOnly = false;
                    //dtOrdernums.Columns["ONMessage"].ReadOnly = false;

                    //Create RMA entry for each Orderno Needed 
                    foreach (DataRow drOrdernums in dtOrdernums.Rows)
                    {
                        SqlString1 = "EXEC hercust..[proc_mag_rma_htinsert] @OrderIncrementid = '" + dJson.order_increment_id
                            + "', @RMAMagentoNumber = '" + dJson.rma.increment_id
                            + "', @RMAEntityid = " + dJson.rma.id.ToString()  //Should all be the same number for all items in RMA
                            + ", @DateRequested = '" + dJson.rma.created_at + "'";

                        if (MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine(SqlString1);
                        }
                        else
                        {
                            dtRMAs = new System.Data.DataTable();
                            dtRMAs = Helper.Sql_Misc_Fetch(SqlString1);     // SHOULD RETURN 1 ROW WITH RMA table ID
                            if (dtRMAs.Rows.Count == 1)
                            {
                                //Ordernums.Add(long.Parse(dtRMAs.Rows[0][0].ToString()));
                                drOrdernums["rmaid"] = long.Parse(dtRMAs.Rows[0][0].ToString());
                            }
                            else
                            {
                                Console.WriteLine(" !! SQL 2 ERROR; ");
                                Returnvalue = false;
                                break;
                            }
                        }
                    }

                    //Create OrderItems Rows in DB 
                    //EXEC Communications..[proc_mag_rma_htinsert_items] 
                    foreach (DataRow drOrdernums in dtOrdernums.Rows)
                    {
                        dtOrderItems = Helper.Sql_Misc_Fetch("SELECT DISTINCT Ordernum, itemnum, ParentItemId "
                            + " FROM herCust..items II WHERE ParentItemId IN (" + sHelper + ") AND Ordernum = " + drOrdernums[0].ToString());

                        foreach (DataRow drOrderItems in dtOrderItems.Rows)
                        {
                            OrderItemReturnsIds = "";
                            for (int zz = 0; zz < dJson.items.Count; zz++)
                            {
                                if (dJson.items[zz].order_item_id == drOrderItems["parentitemid"].ToString() && dJson.items[zz].qty_requested > 0)
                                {
                                    if (MagetnoProductAPI.DevMode > 0)
                                    {
                                        Console.WriteLine("EXEC hercust..[proc_mag_rma_htinsert_orderitemreturns] @OrderItemid = " + dJson.items[zz].order_item_id
                                        + ", @Qty = " + dJson.items[zz].qty_requested
                                         + ", @ReturnType = " + dJson.items[zz].resolution_id
                                        + ", @ReturnReason = " + dJson.items[zz].reason_id
                                        + ", @DateRequested = '" + dJson.rma.created_at + "';");
                                    }
                                    else
                                    {
                                        dtOrderItemIds = new System.Data.DataTable();
                                        dtOrderItemIds = Helper.Sql_Misc_Fetch("EXEC hercust..[proc_mag_rma_htinsert_orderitemreturns] @OrderItemid = " + dJson.items[zz].order_item_id
                                        + ", @Qty = " + dJson.items[zz].qty_requested
                                        + ", @ReturnType = " + dJson.items[zz].resolution_id
                                        + ", @ReturnReason = " + dJson.items[zz].reason_id
                                        + ", @DateRequested = '" + dJson.rma.created_at + "';");

                                        // !! if dJson.items[zz].resolution_id = 1 = exchange; add to string and make exchange info for orderNotes !!

                                        //if (dtOrderItemIds.Rows.Count == 1)
                                        //{
                                        OrderItemReturnsIds += dtOrderItemIds.Rows[0][0].ToString() + ",";

                                        //}
                                    }

                                    //drOrderItems["ONMessage"] = "Qty Returned: " + dJson.items[zz].qty_requested;
                                }
                            }

                            if (OrderItemReturnsIds.Length > 0 && OrderItemReturnsIds.EndsWith(","))
                            {
                                OrderItemReturnsIds = OrderItemReturnsIds.Substring(0, OrderItemReturnsIds.Length - 1);
                            }

                            if (OrderItemReturnsIds.Length > 0)
                            {
                                if (MagetnoProductAPI.DevMode > 0)
                                {
                                    Console.WriteLine("EXEC hercust..[proc_mag_rma_htinsert_map] @ReturnNumber = " + drOrdernums["RmaId"].ToString() + ", @OrderItemReturnsIds = '" + OrderItemReturnsIds + "';");
                                }
                                else
                                {
                                    if (!Helper.Sql_Misc_NonQuery("EXEC hercust..[proc_mag_rma_htinsert_map] @ReturnNumber = " + drOrdernums["RmaId"].ToString() + ", @OrderItemReturnsIds = '" + OrderItemReturnsIds + "';"))
                                    {

                                        Console.WriteLine(" !! SQL 3 ERROR; ");
                                        Returnvalue = false;
                                        break;
                                    }

                                    //foreach (DataRow drOrdernums2 in dtOrdernums.Rows)
                                    //{
                                    //    if 
                                    //    Helper.Sql_Misc_NonQuery("INSERT Hercust..OrderNotes(OrderNo, Posted, Author, Message, IsPublic, IsAction) SELECT " + drOrdernums2["ordernum"].ToString() + ", GETDATE(), 'Magento RMA', 'Magento initiated an RMA: " + Message + "', 0, 0");
                                    //}

                                    //Can Load MAP one at a time if needed
                                    //EXEC hercust..[proc_mag_rma_htinsert_map] @ReturnNumber = 1000, @OrderItemReturnsIds = '', @OrderItemReturnId= 10 
                                }
                            }
                        }

                        // return_label_type : "Pre-paid Mailing Label ($7.50 deducted from refund)" or "Do it yourself"
                        if (MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine("INSERT hercust..Ordernotes(OrderNo, Posted, Author, Message, IsPublic) "
                                + " SELECT " + drOrdernums["ordernum"].ToString() + ", Getdate(), 'Magento RMA', '" + OrderNotesMessage + " for RMA: " + drOrdernums["rmaid"].ToString() + "', 0");
                        }
                        else
                        {
                            Helper.Sql_Misc_NonQuery("INSERT hercust..Ordernotes(OrderNo, Posted, Author, Message, IsPublic) "
                                + " SELECT " + drOrdernums["ordernum"].ToString() + ", Getdate(), 'Magento RMA', '" + OrderNotesMessage + " for RMA: " + drOrdernums["rmaid"].ToString() + "', 0");
                        }
                    }

                    if (MagetnoProductAPI.DevMode == 0)
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE communications..middleware_incoming SET status = 600, worker_id = 777 WHERE ID = " + dr["id"].ToString());
                    }

                }
                catch (Exception ex)
                {
                    //if (MagetnoProductAPI.DevMode > 0)
                    //{
                        Console.WriteLine("ERROR: Process_Incoming_DirectRMA_Insert_3(): " + ex.ToString());
                    //}

                    if (MagetnoProductAPI.DevMode == 0)
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE communications..middleware_incoming SET status = 700, worker_id = 777 WHERE ID = " + dr["id"].ToString());
                    }

                    Returnvalue = false;
                }

            }
            return Returnvalue;
        }

        // !!!!!!!!!!!!! 2024-05-13, 2024-07-05 - USE THIS
        public static Boolean Process_Incoming_DirectRMA_Insert_2(long Middlewareid = 0, int MiddlewareIncomingStatus = 0)
        {
            Boolean bReturnvalue = true;
            System.Data.DataTable dt;
            System.Data.DataTable dtOrdernums;
            System.Data.DataTable dtRMAs;
            System.Data.DataTable dtOrderItems;
            System.Data.DataTable dtOrderItemIds;
            String OrderItemReturnsIds;
            String sHelper;
            List<long> Ordernums = new List<long>();
            String SqlString1;
            dynamic dJson;
            dynamic dJsonItems;

            //load all info to arrays 
            if (Middlewareid > 0)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT id, from_magento FROM Communications..middleware_incoming WHERE id = " + Middlewareid.ToString() + " AND endpoint_method = 'incoming - module RMA' ");
            }
            else
            {
                dt = Helper.Sql_Misc_Fetch("SELECT id, from_magento FROM Communications..middleware_incoming WHERE status = " + MiddlewareIncomingStatus.ToString() + " AND endpoint_method = 'incoming - module RMA' ");
            }

            foreach (DataRow dr in dt.Rows)
            {
                sHelper = "";
                try
                {
                    if (MagetnoProductAPI.DevMode == 0)
                    {
                        if (!Helper.Sql_Misc_NonQuery("UPDATE communications..middleware_incoming SET status = 200, worker_id = 777 WHERE ID = " + dr["id"].ToString()))
                        {
                            Console.WriteLine(" !! SQL 1 ERROR; ");
                            bReturnvalue = false;
                            break;
                        }
                    }

                    dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(dr["from_magento"].ToString());
                    dJsonItems = dJson.rma.items;

                    //Pass 1 for item, see how many orderno are involved for split orders
                    for (int zz = 0; zz < dJsonItems.Count; zz++)
                    {
                        sHelper += dJsonItems[zz].order_item_id;

                        if (zz < dJsonItems.Count-1)
                        {
                            sHelper += ",";
                        }
                    }

                    //if (sHelper.Length > 2)
                    //{
                    //    sHelper = sHelper.Substring(0, sHelper.Length - 2);
                    //}

                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("Items COUNT: " + dJsonItems.Count.ToString());
                        Console.WriteLine("sHelper = " + sHelper + ", dJson.rma.date_requested:" + dJson.rma.date_requested );
                    }

                    dtOrdernums = Helper.Sql_Misc_Fetch("SELECT DISTINCT Ordernum, 0 [RmaId] FROM herCust..items WHERE ParentItemId IN (" + sHelper + ") ORDER BY 1");
                    //dtOrdernums = Helper.Sql_Misc_Fetch("SELECT distinct Ordernum " 
                    //    + ", (SELECT Convert(varchar(32), ParentItemId) + ',' FROM herCust..items IX WHERE IX.ordernum = II.ordernum  AND IX.ParentItemId IN(" + sHelper + ") FOR XML Path('')) [parentitemids] "
                    //    + " FROM herCust..items II WHERE ParentItemId IN (" + sHelper + ")");

                    dtOrdernums.Columns["RmaId"].ReadOnly = false;

                    //Create RMA entry for each Orderno Needed 
                    foreach (DataRow drOrdernums in dtOrdernums.Rows)
                    {
                        SqlString1 = "EXEC hercust..[proc_mag_rma_htinsert] @OrderIncrementid = '" + dJson.rma.order_increment_id  
                            + "', @RMAMagentoNumber = '" + dJson.rma.increment_id
                            + "', @RMAEntityid = " + dJsonItems[0].rma_entity_id.ToString()  //Should all be the same number for all items in RMA
                            + ", @DateRequested = '" + dJson.rma.date_requested + "'";

                        if (MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine(SqlString1);
                        }
                        else
                        {
                            dtRMAs = new System.Data.DataTable();
                            dtRMAs = Helper.Sql_Misc_Fetch(SqlString1);     // SHOULD RETURN 1 ROW WITH RMA table ID
                            if (dtRMAs.Rows.Count == 1)
                            {
                                //Ordernums.Add(long.Parse(dtRMAs.Rows[0][0].ToString()));
                                drOrdernums["rmaid"] = long.Parse(dtRMAs.Rows[0][0].ToString());
                            }
                            else
                            {
                                Console.WriteLine(" !! SQL 2 ERROR; ");
                                bReturnvalue = false;
                                break;
                            }
                        }
                    }

                    //Create OrderItems Rows in DB 
                    //EXEC Communications..[proc_mag_rma_htinsert_items] 
                    foreach (DataRow drOrdernums in dtOrdernums.Rows)
                    {
                        dtOrderItems = Helper.Sql_Misc_Fetch("SELECT DISTINCT Ordernum, itemnum, ParentItemId "
                            + " FROM herCust..items II WHERE ParentItemId IN (" + sHelper + ") AND Ordernum = " + drOrdernums[0].ToString());

                        foreach (DataRow drOrderItems in dtOrderItems.Rows)
                        {
                            OrderItemReturnsIds = "";
                            for (int zz = 0; zz < dJsonItems.Count; zz++)
                            {
                                if (dJsonItems[zz].order_item_id == drOrderItems["parentitemid"].ToString())
                                {
                                    if (MagetnoProductAPI.DevMode > 0)
                                    {
                                        Console.WriteLine("EXEC hercust..[proc_mag_rma_htinsert_orderitemreturns] @OrderItemid = " + dJsonItems[zz].order_item_id
                                        + ", @Qty = " + dJsonItems[zz].qty_approved
                                        + ", @ReturnType = 1"
                                        + ", @DateRequested = '" + dJson.rma.date_requested + "';");
                                    }
                                    else
                                    {
                                        dtOrderItemIds = new System.Data.DataTable();
                                        dtOrderItemIds = Helper.Sql_Misc_Fetch("EXEC hercust..[proc_mag_rma_htinsert_orderitemreturns] @OrderItemid = " + dJsonItems[zz].order_item_id
                                        + ", @Qty = " + dJsonItems[zz].qty_approved
                                        + ", @ReturnType = 1"
                                        + ", @DateRequested = '" + dJson.rma.date_requested + "';");

                                        if (dtOrderItemIds.Rows.Count == 1)
                                        {
                                            OrderItemReturnsIds += dtOrderItemIds.Rows[0][0].ToString() + ",";

                                        }
                                    }
                                }
                            }

                            if (OrderItemReturnsIds.Length > 0)
                            {
                                if (MagetnoProductAPI.DevMode > 0)
                                {
                                    Console.WriteLine("EXEC hercust..[proc_mag_rma_htinsert_map] @ReturnNumber = " + drOrdernums["RmaId"].ToString() + ", @OrderItemReturnsIds = '" + OrderItemReturnsIds + "';");
                                }
                                else
                                {
                                    if (!Helper.Sql_Misc_NonQuery("EXEC hercust..[proc_mag_rma_htinsert_map] @ReturnNumber = " + drOrdernums["RmaId"].ToString() + ", @OrderItemReturnsIds = '" + OrderItemReturnsIds + "';"))
                                    {

                                        Console.WriteLine(" !! SQL 3 ERROR; ");
                                        bReturnvalue = false;
                                        break;
                                    }
                       
                                    //Can Load MAP one at a time if needed
                                    //EXEC hercust..[proc_mag_rma_htinsert_map] @ReturnNumber = 1000, @OrderItemReturnsIds = '', @OrderItemReturnId= 10 
                                }
                            }
                        }
                    }

                    if (MagetnoProductAPI.DevMode == 0)
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE communications..middleware_incoming SET status = 600, worker_id = 777 WHERE ID = " + dr["id"].ToString());
                    }

                }
                catch (Exception ex)
                {
                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("ERROR: Process_Incoming_DirectRMA_Insert_2(): " + ex.ToString());
                    }

                    if (MagetnoProductAPI.DevMode == 0)
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE communications..middleware_incoming SET status = 700, worker_id = 777 WHERE ID = " + dr["id"].ToString());
                    }

                    bReturnvalue = false;
                }

            }
            return bReturnvalue;
        }


        public static Boolean Process_Incoming_DirectRMA_Insert(long Middlewareid = 0, int MiddlewareIncomingStatus=0)
        {
            Boolean bReturnvalue = true;
            System.Data.DataTable dt;
            dynamic dJson;
            dynamic dJsonItems;
            string Returns = "";
            string ReturnReason = "";
            string ReturnCondtion = "";
            string ReturnResolution = "";
            string RMAEntityid = "";

            /*
             *  REWrite:  
             *  1) Loop through json items, get all parerent/magentoids... figure out if split order
             *  2) insert 1 or more RMA rows
             *  3) Insert orderitem rows
             *  4) map them
             * 
             */

            try
            {

                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, from_magento FROM Communications..middleware_incoming WHERE id = " + Middlewareid.ToString() + " AND endpoint_method = 'incoming - module RMA' ");
                }
                else
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, from_magento FROM Communications..middleware_incoming WHERE status = " + MiddlewareIncomingStatus.ToString() + " AND endpoint_method = 'incoming - module RMA' ");
                }

                foreach (DataRow dr in dt.Rows)
                {
                    dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(dr["from_magento"].ToString());

                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("dJson: " + dJson.rma);
                    }

                    dJsonItems = dJson.rma.items;

                    if (MagetnoProductAPI.DevMode > 1)
                    {
                        Console.WriteLine("Items COUNT: " + dJsonItems.Count.ToString());
                    }

                    for (int zz = 0; zz < dJsonItems.Count; zz++)
                    {
                        if (MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine("Items: " + dJsonItems[zz].ToString());
                        }
                        Returns += dJsonItems[zz].order_item_id + "^" + dJsonItems[zz].qty_authorized + ",";
                        ReturnReason += dJsonItems[zz].order_item_id + "^" + dJsonItems[zz].reason + ",";
                        ReturnCondtion += dJsonItems[zz].order_item_id + "^" + dJsonItems[zz].condition + ",";
                        ReturnResolution += dJsonItems[zz].order_item_id + "^" + dJsonItems[zz].resolution + ",";

                        RMAEntityid = dJsonItems[zz].rma_entity_id;
                    }

                    if (MagetnoProductAPI.DevMode > 0) // 1  !!
                    {
                        Console.WriteLine(" -- Returns: " + Returns + "; RMAEntityid: " + RMAEntityid);
                        Console.WriteLine(" --------------------- ");

                        Console.WriteLine("SQL:  EXEC HERCUST..proc_mag_RMA_HTInsert "
                            + " @OrderIncrementid = " + dJson.order_increment_id
                            + " , @Returns = '" + Returns + "' "
                            + " , @RMAMagentoNumber = '" + dJson.increment_id + "' "
                            + " , @RMAEntityid = " + RMAEntityid
                            + " , @DateRequested = '" + dJson.date_requested + "' "
                            + " , @Reason = " + ReturnReason
                            + " , @Condition = " + ReturnCondtion
                            + " , @Resolution = " + ReturnResolution);
                    }
                    else
                    {
                        Helper.Sql_Misc_NonQuery("EXEC HERCUST..proc_mag_RMA_HTInsert "
                            + " @OrderIncrementid = " + dJson.order_increment_id
                            + " , @Returns = '" + Returns + "' "
                            + " , @RMAMagentoNumber = '" + dJson.increment_id + "' "
                            + " , @RMAEntityid = 1520 "
                            + " , @DateRequested = '" + dJson.date_requested + "' "
                            + " , @Reason = " + ReturnReason
                            + " , @Condition = " + ReturnCondtion
                            + " , @Resolution = " + ReturnResolution);
                    }

                }
            }
            catch (Exception ex)
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                  Console.WriteLine("ERROR Process_Incoming_DirectRMA_Insert(): " + ex.ToString());
                }
                bReturnvalue = false;
            }

            return bReturnvalue;
        }

        public static Boolean Process_Incoming_RMA(long Middlewareid)  // =0 later, make mandatory for testing
        {
            Boolean bReturnvalue = true;
            System.Data.DataTable dt = new System.Data.DataTable();
            String APIReturn;
            dynamic dJson;
            dynamic dJsonItems;
            Magento_MCP.MagentoModels.SalesModels.Rma RMA;
            List<Magento_MCP.MagentoModels.HelperModels.RmaItem> RMAItems;
            Magento_MCP.MagentoModels.HelperModels.RmaItem RMAItem;
            //Magento_MCP.MagentoModels.SalesModels.RmaDataObject RMDdo;
            string MagentoOrderno;
            string MagentoItemIds;
            System.Data.DataTable RMAExistCheckFound;

            if (Middlewareid > 0)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT id, status, from_magento FROM Communications..Middleware WHERE source_table = 'NEW RMA' AND ID = " + Middlewareid.ToString());
            }

            foreach (DataRow dr in dt.Rows)
            {
                if (dr["status"].ToString() == "100")
                {
                    APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");
                }

                if (dr["status"].ToString() == "600")
                {
                    dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(dr["from_magento"].ToString());

                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("COUNT: " + dJson.Count.ToString());
                    }

                    for (int xx = 0; xx < dJson.Count; xx++)
                    {
                        if (MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine("order_incrementid: " + dJson[xx].order_increment_id);
                        }

                        MagentoOrderno = dJson[xx].order_increment_id;
                        MagentoItemIds = "";

                        RMAItems = new List<Magento_MCP.MagentoModels.HelperModels.RmaItem>();
                        RMA = new Magento_MCP.MagentoModels.SalesModels.Rma();
                        RMA.RmaDataObject = new Magento_MCP.MagentoModels.SalesModels.RmaDataObject();

                        RMA.RmaDataObject.IncrementId = dJson[xx].rma.increment_id;
                        RMA.RmaDataObject.OrderIncrementId = dJson[xx].order_increment_id;
                        RMA.RmaDataObject.OrderId = dJson[xx].order_id;

                        RMA.RmaDataObject.StoreId = 2;
                        if (dJson[xx].rma.increment_id.ToString().Substring(0,1) == "0")
                        {
                            RMA.RmaDataObject.StoreId = 2;
                        }

                        RMA.RmaDataObject.DateRequested = dJson[xx].rma.created_at;
                        RMA.RmaDataObject.CustomerCustomEmail = "";
                        RMA.RmaDataObject.Status = "authorized";

                        dJsonItems = dJson[xx].items;

                        if (MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine("Items COUNT: " + dJsonItems.Count.ToString());
                        }

                        for(int zz=0; zz < dJsonItems.Count; zz++)
                        {
                            RMAItem = new Magento_MCP.MagentoModels.HelperModels.RmaItem();

                            MagentoItemIds += dJsonItems[zz].order_item_id + ",";

                            if (MagetnoProductAPI.DevMode > 0)
                            {
                                Console.WriteLine("Items zz: " + dJsonItems[zz].order_item_id);
                            }

                            if (dJsonItems[zz].qty_requested > 0)
                            {
                                //RMAItem.EntityId = dJsonItems.order_item_id;
                                RMAItem.RmaEntityId = dJsonItems[zz].rma_id;
                                RMAItem.OrderItemId = dJsonItems[zz].order_item_id;
                                RMAItem.QtyRequested = dJsonItems[zz].qty_requested;
                                RMAItem.QtyAuthorized = dJsonItems[zz].qty_requested;
                                RMAItem.QtyApproved = dJsonItems[zz].qty_requested;
                                RMAItem.QtyReturned = dJsonItems[zz].qty_requested;
                                RMAItem.Reason = dJsonItems[zz].reason_id;
                                RMAItem.Condition = dJsonItems[zz].condition_id;
                                RMAItem.Resolution = dJsonItems[zz].resolution_id;
                                RMAItem.Status = "authorized";      //harded for now, not really used in MCP RMA

                                //RMA.RmaDataObject.Items = new List<Magento_MCP.MagentoModels.HelperModels.RmaItem>;
                                RMAItems.Add(RMAItem);
                            }
                        }

                         RMA.RmaDataObject.Items = RMAItems;

                        var TM_JsonChild = JsonConvert.SerializeObject(RMA, Formatting.Indented, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
                        TM_JsonChild = TM_JsonChild.Replace("rmaDataObject", "rma");
                        TM_JsonChild = TM_JsonChild.Replace("'", "`");

                        if (MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine(TM_JsonChild);
                        }

                        //CHECK FOR DUPLIDATE ENTRIES
                        RMAExistCheckFound =  Helper.Sql_Misc_Fetch("EXEC Communications..proc_mag_rma_exising_check @MagentoOrderno = '" + MagentoOrderno + "', @MagentoItemIds = '" + MagentoItemIds + "'; ");
                        if (RMAExistCheckFound.Rows.Count > 0)
                        {
                            if (RMAExistCheckFound.Rows[0]["found"].ToString() == "0")
                            {
                                //2024-03-26: Added source_id to insert
                                //Helper.Sql_Misc_NonQuery("INSERT communications..MIDDLEWARE_INCOMING(posted, status, worker_id, from_magento, endpoint_name, endpoint_method) " +
                                //    " SELECT Getdate(), 0, 102, '" + TM_JsonChild + "', 'INCOMING', 'INCOMING - Module RMA'; ");

                                //DEPLOYED 2024-03-26 4:35PM 
                                //Helper.Sql_Misc_NonQuery("INSERT communications..MIDDLEWARE_INCOMING(posted, status, worker_id, from_magento, endpoint_name, endpoint_method, source_id) " +
                                //    " SELECT Getdate(), 0, 103, '" + TM_JsonChild + "', 'INCOMING', 'INCOMING - Module RMA', '" + RMA.RmaDataObject.IncrementId + "' ");

                                Helper.Sql_Misc_NonQuery("IF 0 = (SELECT COUNT(*) FROM middleware_incoming WHERE source_id = '" + RMA.RmaDataObject.IncrementId + "' AND endpoint_method = 'INCOMING - Module RMA') "
                                    + " INSERT communications..MIDDLEWARE_INCOMING(posted, status, worker_id, from_magento, endpoint_name, endpoint_method, source_id) " 
                                    + " SELECT Getdate(), 0, 105, '" + TM_JsonChild + "', 'INCOMING', 'INCOMING - Module RMA', '" + RMA.RmaDataObject.IncrementId + "' ");

                           }
                        }
                        else
                        {
                            if (MagetnoProductAPI.DevMode > 0)
                            {
                                Console.WriteLine("NOT FOUND RMAExistCheckFound.Rows==0; @MagentoOrderno = '" + MagentoOrderno);
                            }
                        }
                    }
                }
            }

            return bReturnvalue;
        }


        //2024-03-26
        public static Boolean ReturnLabelPDF_API_Process()
        {
            Boolean Returnvalue = true;

            System.Data.DataTable dtRMApi;
            dtRMApi = Helper.Sql_Misc_Fetch("SELECT TOP 25 id, RMANumber FROM communications..RMA_Label_Emails WHERE ISNULL(Status, 100) = 100 AND LEN(ISNULL(RmaNumber,'')) > 0");
            foreach (DataRow drRMApi in dtRMApi.Rows)
            {
                ReturnLabelPDF_API(drRMApi["RMANumber"].ToString());
            }
            return Returnvalue;
        }

        public static Boolean ReturnLabelPDF_API(string RMANumber)
        {
            Boolean Returnvalue = true;
            rmaPost RMAPost = new rmaPost();
            String APIResponse = "";
            string RMAShippingLabelAPI;

            try
            {
                if (MagetnoProductAPI.DevMode > 1)
                {
                    RMAShippingLabelAPI = ConfigurationManager.AppSettings["RMAShippingLabelAPITEST"];
                }
                else
                {
                    RMAShippingLabelAPI = ConfigurationManager.AppSettings["RMAShippingLabelAPI"];
                }

                var content = new StringContent(@"{""RMANumber"":""" + RMANumber + @"""}", System.Text.Encoding.UTF8, "application/json");
                var url = new Uri(ConfigurationManager.AppSettings["GetRMAPDFLinkforMagento"]);

                string posting = Helper.POSTData(content, url);
                posting = posting.Substring(6);
                posting = posting.Substring(0, posting.Length-2);
                if (MagetnoProductAPI.DevMode > 1)
                {
                    Console.WriteLine(posting);
                }
               
                if (posting.Length > 50)
                {
                    RMAPost.RMAShippingLabel = new RMAShippingLabel();

                    RMAPost.RMAShippingLabel.increment_id = RMANumber;
                    RMAPost.RMAShippingLabel.status_id = "2";
                    RMAPost.RMAShippingLabel.pdf_attachment = posting;

                    /*
	                    1: Pending Approval
                        2: Approved
                        3: Rejected
                        4: Package Sent
                        5: Closed (edited) 
                     */


                    dynamic dJson;
                    dJson = Newtonsoft.Json.JsonConvert.SerializeObject(RMAPost);

                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine(" --------------------- ");
                        Console.WriteLine(dJson.ToString());
                        Console.WriteLine(" --------------------- ");
                    }

                    APIResponse = Helper.MagentoApiPush_Direct("POST", RMAShippingLabelAPI, dJson.ToString());  
                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine(APIResponse);
                    }

                    if (APIResponse.ToLower() != "ok")
                    {
                        Console.WriteLine("ERROR");
                        Helper.Sql_Misc_NonQuery("UPDATE communications..RMA_Label_Emails SET status = CASE WHEN Status = 700 THEN 701 ELSE 700 END WHERE RMANumber = '" + RMANumber + "' ;");

                        Returnvalue = false;
                    }
                    else
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE communications..RMA_Label_Emails SET status = 600 WHERE RMANumber = '" + RMANumber + "' ;");
                    }
                }
            }
            catch (Exception exx)
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine(exx.ToString());
                }
                Returnvalue = false;
            }

            return Returnvalue;
        }

        public static Boolean Process_RMA_Status_Update(long Middlewareid = 0)
        {
            Boolean ReturnValue = true;
            System.Data.DataTable dt = new System.Data.DataTable();
            String APIReturn = "";

            try
            {
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE source_table = 'hercust..rmaupdate' AND id = " + Middlewareid.ToString());
                }
                else
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE Status=100 AND source_table = 'hercust..rmaupdate' ORDER BY ID");
                }

                Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                {
                    APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");

                    if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                    {
                        //Middleware should already by set to 500
                        if (MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine("ERROR Process_Order_Invoice: " + dr["source_id"].ToString());
                        }

                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, from_magento = 'ERROR... " + APIReturn + "' WHERE ID = " + dr["id"].ToString());

                        Orders.Order_AddHTNote("Middleware", "API RMA Update Failed, status not updated in Magento; please do so manually.", long.Parse(dr["source_id"].ToString()), false, false);
                        ReturnValue = false;
                    }
                    else //Should return new value
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_Magento = 'done' WHERE ID = " + dr["id"].ToString());

                        Orders.Order_AddHTNote("Middleware", "API RMA Update Successful.", long.Parse(dr["source_id"].ToString()), false, false);
                    }

                }
            }
            catch (Exception ex)
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine("ERROR Process_RMA_Status_Update: " + ex.ToString());
                }
                ReturnValue = false;
            }

            return ReturnValue;
        }

    }

    class Customers
    {
        public static Boolean Customer_Bulk_Process(int maxtorun = 100, string Environment = "")
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt = Helper.Sql_Misc_Fetch("SELECT TOP " + maxtorun.ToString() + " id FROM Communications..middleware with (nolock) WHERE status = 100 AND source_table = 'hercust..cust' ORDER BY ID ");
            foreach (DataRow dr in dt.Rows)
            {
                Process_Customer_toMagento_Middleware(long.Parse(dr["id"].ToString()), Environment);
            }

            return true;
        }
      

        //Middleware Rows to be Created in MCP FOR NOW
        public static Boolean Process_Customer_toMagento_Middleware(long Middlewareid, string Environment = "")
        {
            Boolean bReturnvalue = true;
            System.Data.DataTable dt = new System.Data.DataTable();
            String APIReturn = "";
            long mid;

            try
            {
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND source_table = 'hercust..cust' AND ISNULL(Batchnum,0) = 0 AND id = " + Middlewareid.ToString());
                }
                //else
                //{}
                
                Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                {
                    APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString()), Environment).Replace("\"", "");

                    if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                    {
                        //Middleware should already by set to 500
                        Console.WriteLine("ERROR for Process_Product_Configurable_Middleware Processing: " + dr["source_id"].ToString());
                    }
                    else //Should return new value
                    {
                        if (long.TryParse(APIReturn, out mid) && dr["endpoint_method"].ToString().ToUpper() == "POST")
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + mid.ToString().Replace("'", "''") + "' WHERE ID = " + dr["id"].ToString());
                            Helper.Sql_Misc_NonQuery("UPDATE Hercust..Cust SET Magentoid = " + mid.ToString() + " WHERE custno = '" + dr["source_id"].ToString() + "' AND Magentoid IS NULL");
                        }
                        else
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, error_message='No MID Returned' WHERE ID = " + Middlewareid.ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Process_Customer_toMagento_Middleware(" + Middlewareid.ToString() + ": " + ex.ToString());
                bReturnvalue = false;
                Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, error_message='ERROR: " + ex.ToString() + "' WHERE ID = " + Middlewareid.ToString());
            }

            return bReturnvalue;
        }

        //2024-10-14
        //LastEntityid == -1 GET LAST ENTITY IN DB AND DO ALL
        public static Boolean Contactus_Fetch_Process(long Entityid = 0, string Email = "", long LastEntityid = 0)
        {
            Boolean Returnvalue = true;
            String APIReturn = "";
            System.Data.DataTable dt;
            System.Data.DataTable dtLatest;
            dynamic dJson;
            dynamic dJsonItems;
            string SqlString;

            try
            {
                if (Entityid > 0)
                {
                    APIReturn = Helper.MagentoApiGet("/V1/all/contact_us/?searchCriteria[filter_groups][0][filters][0][field]=entity_id&searchCriteria[filter_groups][0][filters][0][value]=" + Entityid.ToString() + "&searchCriteria[filter_groups][0][filters][0][condition_type]=eq");
                }
                else if (LastEntityid > 0)
                {
                    APIReturn = Helper.MagentoApiGet("/V1/all/contact_us/?searchCriteria[filter_groups][0][filters][0][field]=entity_id&searchCriteria[filter_groups][0][filters][0][value]=" + LastEntityid.ToString() + "&searchCriteria[filter_groups][0][filters][0][condition_type]=gt");
                }
                else if (Email != "")
                {
                    APIReturn = Helper.MagentoApiGet("/V1/all/contact_us/?searchCriteria[filter_groups][0][filters][0][field]=" + LastEntityid.ToString() + "&searchCriteria[filter_groups][0][filters][0][value]=" + Email + "&searchCriteria[filter_groups][0][filters][0][condition_type]=eq");
                }
                else if (Entityid == 0 && LastEntityid == -1)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT ISNULL(Max(magentoid_contactus),0) [mid] FROM communications..magento_contactus");

                    if (dt.Rows[0]["mid"].ToString() != "0")
                    {
                        APIReturn = Helper.MagentoApiGet("/V1/all/contact_us/?searchCriteria[filter_groups][0][filters][0][field]=entity_id&searchCriteria[filter_groups][0][filters][0][value]=" + dt.Rows[0]["mid"].ToString() + "&searchCriteria[filter_groups][0][filters][0][condition_type]=gt");
                    }
                }

                dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(APIReturn);
                dJsonItems = dJson.items;

                for (int xx = 0; xx < dJsonItems.Count; xx++)
                {
                    SqlString = "EXEC communications..proc_mag_get_send_contactus_table " +
                       " @MiddlewareIncomingID = 0" +
                       ", @CustomerName = '" + dJsonItems[xx].name.ToString().Replace("'", "''") + "' " +
                       ", @Email = '" + dJsonItems[xx].email + "' " +
                       ", @OrderNumber = '" + dJsonItems[xx].order + "' " +
                       ", @Subject = '" + dJsonItems[xx].subject.ToString().Replace("'", "''") + "' " +
                       ", @Brand = '" + dJsonItems[xx].brand.ToString().Replace("'", "''") + "' " +
                       ", @Band_size = '" + dJsonItems[xx].band_size.ToString().Replace("'", "''") + "' " +
                       ", @Cup_size = '" + dJsonItems[xx].cup_size.ToString().Replace("'", "''") + "' " +
                       ", @Comment = '" + dJsonItems[xx].comment.ToString().Replace("'", "''") + "' " +     //+ json.form.comment + "' " +
                       ", @form_type = 'Contact Us Form' " +
                       ", @store_id = '" + dJsonItems[xx].store_id + "' " +     // json.form.store_id + "' ";
                       ", @BrowserInfo = '" + dJsonItems[xx].user_agent + "; " + dJsonItems[xx].ip_address + "; " + dJsonItems[xx].browser_name + "; " + dJsonItems[xx].browser_version + "' " +
                       ", @Magentoid = " + dJsonItems[xx].entity_id;

                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine(SqlString);
                    }

                    if (MagetnoProductAPI.DevMode < 2)
                    {
                        Helper.Sql_Misc_NonQuery(SqlString);

                        dtLatest = Helper.Sql_Misc_Fetch("SELECT TOP 1 ISNULL(ID,0) [id] FROM communications..Magento_Contactus where KustomerEmail = '" + dJsonItems[xx].email + "' ORDER BY ID Desc ");
                        if (dtLatest.Rows.Count == 1 && dtLatest.Rows[0]["id"].ToString() != "0")
                        {
                            Kustomer.PushKustomerOrderNotes(long.Parse(dtLatest.Rows[0]["id"].ToString()));

                            //Send EMail to CSRs - 2024-10-22
                            //HisRoom not getting to Kustomer, push from here and send CSR Emails - 2024-10-28
                            if (dJsonItems[xx].subject.ToString().ToLower().Substring(0,8) == "hisroom")
                            {
                                Kustomer.Process_Kustomer_Contactus(int.Parse(dtLatest.Rows[0]["id"].ToString()), 1, true, true);
                            }
                            else
                            {
                                Kustomer.Process_Kustomer_Contactus(int.Parse(dtLatest.Rows[0]["id"].ToString()), 1, true, false);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine("ERROR: " + ex.ToString());
                }
                Returnvalue = false;
            }     

            return Returnvalue;
        }

        //2024-10-14
        public static Boolean PDPQuestion_Fetch_Process(long Entityid = 0, long LastEntityid = 0)
        {
            Boolean Returnvalue = true;
            String APIReturn = "";
            System.Data.DataTable dt;
            System.Data.DataTable dtLatest;
            dynamic dJson;
            dynamic dJsonItems;
            string SqlString;

            try
            {
                if (Entityid > 0)
                {
                    APIReturn = Helper.MagentoApiGet("/V1/all/questions/?searchCriteria[filter_groups][0][filters][0][field]=entity_id&searchCriteria[filter_groups][0][filters][0][value]=" + Entityid.ToString() + "&searchCriteria[filter_groups][0][filters][0][condition_type]=eq");
                }
                else if (LastEntityid > 0)
                {
                    APIReturn = Helper.MagentoApiGet("/V1/all/questions/?searchCriteria[filter_groups][0][filters][0][field]=entity_id&searchCriteria[filter_groups][0][filters][0][value]=" + LastEntityid.ToString() + "&searchCriteria[filter_groups][0][filters][0][condition_type]=gt");
                }
                else if (LastEntityid == -1)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT ISNULL(Max(magentoid_question),0) [mid] FROM communications..magento_contactus");

                    if (dt.Rows[0]["mid"].ToString() != "0")
                    {
                        if (MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine("MID: " + dt.Rows[0]["mid"].ToString());
                        }
                        APIReturn = Helper.MagentoApiGet("/V1/all/questions/?searchCriteria[filter_groups][0][filters][0][field]=entity_id&searchCriteria[filter_groups][0][filters][0][value]=" + dt.Rows[0]["mid"].ToString() + "&searchCriteria[filter_groups][0][filters][0][condition_type]=gt");
                    }
                }

                if (MagetnoProductAPI.DevMode > 1)
                {
                    Console.WriteLine(APIReturn);
                }

                dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(APIReturn);
                dJsonItems = dJson.items;

                for (int xx = 0; xx < dJsonItems.Count; xx++)
                {
                    SqlString = "EXEC communications..proc_mag_get_send_pdp_question_table " +
                        " @MiddlewareIncomingID = 0" +
                        ", @CustomerName = '" + dJsonItems[xx].customer_name.ToString().Replace("'", "''") + "' " +
                        ", @Email = '" + dJsonItems[xx].email.ToString().Replace("'", "''") + "' " +
                        ", @Content = '" + dJsonItems[xx].content.ToString().Replace("'", "''") + "' " +
                        ", @url = '" + dJsonItems[xx].url.ToString().Replace("'", "''") + "' " +
                        ", @sku = '" + dJsonItems[xx].sku.ToString().Replace("'", "''") + "' " +
                        ", @Brand = '" + dJsonItems[xx].brand.ToString().Replace("'", "''") + "' " +
                        ", @ToEmail ='" + ConfigurationManager.AppSettings["CSREmail"].ToString() + "' " +
                        ", @store_id = '" + dJsonItems[xx].store_id + "' " +
                        ", @Sizedata = '" + dJsonItems[xx].size_data.ToString().Replace("'", "''") + "' " +
                        ", @BrowserInfo = '" + dJsonItems[xx].user_agent + "; " + dJsonItems[xx].ip_address + "; " + dJsonItems[xx].browser_name + "; " + dJsonItems[xx].browser_version + "' " +
                        ", @Magentoid = " + dJsonItems[xx].entity_id;

                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine(SqlString);
                    }

                    if (MagetnoProductAPI.DevMode < 2)
                    {
                        Helper.Sql_Misc_NonQuery(SqlString);

                        dtLatest = Helper.Sql_Misc_Fetch("SELECT TOP 1 ISNULL(ID,0) [id] FROM communications..Magento_Contactus where KustomerEmail = '" + dJsonItems[xx].email + "' ORDER BY ID Desc ");
                        if (dtLatest.Rows.Count == 1 && dtLatest.Rows[0]["id"].ToString() != "0")
                        {
                            Kustomer.PushKustomerOrderNotes(long.Parse(dtLatest.Rows[0]["id"].ToString()));

                            //Send EMail to CSRs - 2024-10-22
                            Kustomer.Process_Kustomer_Contactus(int.Parse(dtLatest.Rows[0]["id"].ToString()), 1, true, false);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine("ERROR: " + ex.ToString());
                }
                Returnvalue = false;
            }

            return Returnvalue;
        }

    }

    class POHerTools
    {
        /*
        public static Boolean PO_HT_Process_Depricated(long Middlewareid = 0)
        {
            Boolean ReturnValue = true;
            System.Data.DataTable dt;
            System.Data.DataTable dtSku;
            DataSet ds;
            string PONumber;
            string SqlString;
            int LineNumber = 1;
            string POSql;
            System.Data.DataTable dtMfr;
            System.Data.DataTable dtCost;

            if (Middlewareid > 0)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT * FROM communications..Middleware WHERE Source_table = 'RUNPOAPP' AND ID = " + Middlewareid.ToString());
            }
            else
            {
                dt = Helper.Sql_Misc_Fetch("SELECT * FROM communications..Middleware WHERE status = 100 AND Source_table = 'RUNPOAPP' ");
            }

            Helper.Middleware_Status_Update(dt, 200);

            foreach (System.Data.DataRow dr in dt.Rows)
            {
                PONumber = "";
                PONumber = dr["to_magento"].ToString().Remove(0, 9);  // IE: PONUMBER: is in Middleware

                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine("PONumber: " + PONumber + "; Mfr: " + dr["source_id"].ToString());
                }

                if (POCreate(dr["source_id"].ToString(), PONumber))
                {
                    //Datatables: 1-Sku info; 2-Mfr info
                    ds = Helper.Sql_Misc_Fetch_Dataset("EXEC HerRoom..[proc_po_process_json] @PONumber = '" + PONumber + "', @MFRCode = '" + dr["source_id"].ToString() + "'");

                    //Insert Hercust..PurchaseOrder Row
                    dtMfr = ds.Tables[1];
                    dtCost = ds.Tables[2];
                    POSql = "INSERT INTO HerCust..PurchaseOrders(PoNum, Revision, IssueDate, StartDate, CancelDate, AutoCancel, IsFashion, Pieces, OrderValue, Lines) "
                          + " SELECT '" + dtMfr.Rows[0]["ponumber"].ToString() + "' "
                          + ", 0"
                          + ", Getdate()"
                          + ", '" + dtMfr.Rows[0]["postartshipdate1"].ToString() + "' "
                          + ", '" + dtMfr.Rows[0]["pocanceldate1"].ToString() + "' "
                          + ", '" + dtMfr.Rows[0]["poisfashion"].ToString() + "' "
                          + ", '" + dtCost.Rows[0]["pieces"].ToString() + "' "
                          + ", '" + dtCost.Rows[0]["cost"].ToString() + "' "
                          + ", '" + dtCost.Rows[0]["lines"].ToString() + "' "
                          + " WHERE 0 = (SELECT COUNT(*) FROM Hercust..PurchaseOrders WHERE PONum = '" + dtMfr.Rows[0]["ponumber"].ToString() + "')";

                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("POSql: " + POSql);
                        Console.WriteLine(" ----------- ");
                    }
                    else
                    {
                        Helper.Sql_Misc_NonQuery(POSql);

                        // update manufacturer record
                        Helper.Sql_Misc_NonQuery("UPDATE herroom..manufacturers SET DateLastPO=getdate(),POacct1=NULL,POshipvia1=NULL,POterms1=NULL,POnumber=NULL,POnote=NULL,POStartShipDate1=NULL,POCancelDate1=NULL,POAutoCancel=1,POIsFashion=0 WHERE ManufacturerCode='" + dtMfr.Rows[0]["ManufacturerCode"].ToString() + "'");
                    }

                    // looop: 
                    if (ds.Tables.Count > 1)
                    {
                        dtSku = ds.Tables[0];

                        foreach (DataRow drSku in dtSku.Rows)
                        {
                            //add to expected items table

                            SqlString = "INSERT INTO hercust..expected (UPC, NumberExp, UnitCost, Notes, DateIn, PONumber, LineNum)"
                             + " VALUES ('" + drSku["upc"].ToString() + "', " + drSku["poOrder"].ToString() + "," + drSku["ourcost"].ToString() + ", "
                             + "'From NewTools', Getdate(),'" + PONumber + "', " + LineNumber.ToString() + ");"
                             + "INSERT INTO hercust..PoArchive (UPC, NumberExp, UnitCost, DateIn, PONumber)"
                             + " VALUES ('" + drSku["upc"].ToString() + "', " + drSku["poOrder"].ToString() + ", " + drSku["ourcost"].ToString()
                             + ", getdate(),'" + PONumber + "')";

                            if (MagetnoProductAPI.DevMode > 0 || MagetnoProductAPI.DevMode == -1)
                            {
                                Console.WriteLine(SqlString);
                                Console.WriteLine("UPDATE Herroom..Items SET POOrder = 0, POcost=NULL, POcostPerm=0 WHERE UPC = '" + drSku["upc"].ToString() + "'");
                                Console.WriteLine(" ------------ ");
                            }
                            else
                            {
                                //run it

                                Helper.Sql_Misc_NonQuery(SqlString);
                                Helper.Sql_Misc_NonQuery("UPDATE Herroom..Items SET POOrder = 0, POcost=NULL, POcostPerm=0 WHERE UPC = '" + drSku["upc"].ToString() + "'");
                            }
                            LineNumber++;
                        }
                    }
                }
            }

            return ReturnValue;
        }
        */

        //Devmode=2:: don't run, Devmode=1:: run with output, Devmode=0:: run with no output
        public static Boolean PO_HT_Process(long Middlewareid = 0)
        {
            Boolean ReturnValue = true;
            System.Data.DataTable dt;
            System.Data.DataTable dtMfr;
            string PONumber;

            if (Middlewareid > 0)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT * FROM communications..Middleware WHERE Source_table = 'RUNPOAPP' AND ID = " + Middlewareid.ToString());
            }
            else
            {
                dt = Helper.Sql_Misc_Fetch("SELECT * FROM communications..Middleware WHERE status = 100 AND Source_table = 'RUNPOAPP' ");
            }

            Helper.Middleware_Status_Update(dt, 200);

            foreach (System.Data.DataRow dr in dt.Rows)
            {
                PONumber = "";
                PONumber = dr["to_magento"].ToString().Remove(0, 9);  // IE: PONUMBER: is in Middleware

                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine("PONumber: " + PONumber + "; Mfr: " + dr["source_id"].ToString());
                }

                //Insert Hercust..PurchaseOrders Row at this Point, will be used later :: 2024-10-16
                Helper.Sql_Misc_NonQuery("EXEC HerRoom..[proc_po_purchaseorders_insert] @MFRCode = '" + dr["source_id"].ToString() + "'");

                if (POCreate(dr["source_id"].ToString(), PONumber))
                {
                    /*
                    // USE THIS IF WE NEED TO EMAIL THE PDF/Labels to Buyers instead of just copying the file to k::\
                    dtMfr = Helper.Sql_Misc_Fetch("SELECT ISNULL(SendLabels,0) [SendLabels] "
                    + ", CASE WHEN ISNULL(BF.Email, '') = '' AND ISNULL(BH.Email, '') = '' THEN 'buyers@andragroup.com' "
                    + " WHEN BF.id = BH.id AND ISNULL(BH.Email, '') <> '' THEN ISNULL(BH.Email, 'buyers@andragroup.com') "
                    + " WHEN BF.id <> BH.id THEN ISNULL(BF.Email, 'buyers@andragroup.com') + ';' + ISNULL(BH.Email, 'buyers@andragroup.com') "
                    + " ELSE 'buyers@andragroup.com' END [email] "
                    + " FROM herroom..Manufacturers MM "
                    + " LEFT OUTER JOIN Herroom..Buyers BF ON BF.id = MM.BuyerID "
                    + " LEFT OUTER JOIN Herroom..Buyers BH ON BH.id = MM.BuyerHisID WHERE ManufacturerCode = '" + dr["source_id"].ToString() + "'");
                    */

                    dtMfr = Helper.Sql_Misc_Fetch("SELECT ISNULL(SendLabels,0) [SendLabels] "
                     + " FROM herroom..Manufacturers MM WHERE ManufacturerCode = '" + dr["source_id"].ToString() + "'");

                    if (MagetnoProductAPI.DevMode >= 2)
                    {
                        Console.WriteLine("EXEC HerRoom..[proc_po_process_json_expected_archive_inserts] @MFRCode = '" + dr["source_id"].ToString() + "', @Debugging=1");
                    }

                    else // (MagetnoProductAPI.DevMode == 0)
                    {
                        ReturnValue = Helper.Sql_Misc_NonQuery("EXEC HerRoom..[proc_po_process_json_expected_archive_inserts] @MFRCode = '" + dr["source_id"].ToString() + "'; ");

                        if (MagetnoProductAPI.DevMode >= 1)
                        {
                            Console.WriteLine("PO_HT_Process INSERTS: " + ReturnValue.ToString());
                        }

                        //Retrieve PDF SKU Labels if generated; get last file, move all files to archive
                        if (dtMfr.Rows[0]["SendLabels"].ToString() == "1")
                        {
                            //There should be a file on \\10.10.1.113\e$\www\barcodeGen3\PDF_Files\--MFRCode--
                            //Copy Lastest file to k:\
                            //2024-08-09 COMPLETE/TEST !!!
                            PO_PDFLabels_Send(dr["source_id"].ToString(), PONumber); //, dtMfr.Rows[0]["Email"].ToString());  // FINISH - have it grab and move files and meail to buyers
                        }

                        if (ReturnValue)
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET Status = 600 WHERE ID = " + dr["id"].ToString());

                            //Send PO Emails (sproc check mfr..RequestEdiConfirm == 1)
                            Helper.Sql_Misc_NonQuery("EXEC herroom..[proc_po_sendediconfirmation_email] @PoNumber = '" + PONumber + "', @Mfrcode ='" + dr["source_id"].ToString() + "'");
                        }
                        else
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET Status = 700, Error_Message='ERROR INSERTING SQL' WHERE ID = " + dr["id"].ToString());
                        }
                    }
                }
                else
                {
                    Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET Status = 700, Error_Message='POCREATE() ERROR' WHERE ID = " + dr["id"].ToString());
                }
            }

            return ReturnValue;
        }

        //ONLY HAVE MCP Create middleware row AFTER Dropship order invoices, send alerts otherwise !!
        //      it should invoice IMMEDIATELY
        // Process_DropShipPOs(): Select * from Communications..Middleware WHERE Status = 100 AND Source_Table = 'DropshipPO'... source_id is Orderno/po Number
        // Need to rewrite this for Order POs: EXEC HerRoom..[proc_po_process_json] @PONumber = ... to get Hercust..ORders/items QTY where Dropship=1, etc
        // DON'T need: PO_PDFLabels_Send

        public static Boolean POCreate(string MFRCode, string PONumber, Boolean IncludeOptionalCode = false)
        {
            Boolean ReturnValue = true;
            string OutFilename;
            string FilenameXls;
            string OutputLine;
    
            HelperModels.POEDI.Address Addr;
            HelperModels.POEDI.Notes note;
            HelperModels.POEDI.Terms term;
            HelperModels.POEDI.ProductorItemDescriptions ProductorItemDescription;

            HelperModels.POEDI.FOBRelatedInstruction FOBRelatedInstr;
            HelperModels.POEDI.LineItem[] LIs;

            //OPTIONAL FIELDS: BOOLEAN IncludeOptionalCode 
            // FOBRelatedInstruction  - check
            // CarrierInformation  ???
            // Header - Notes  - check
            // LineItem - Date  --  ???
            // Summary   - check
            // items.ExtendedItemTotal	ExtendedItemTotal - check  
            // orderheader - CustomerOrderNumber
            // address - addressname  - check
            // address - country - check 

            // HerRoom..[proc_po_process_json] will get info for this

            /*
                ASP CODE TO CALL .NET EXE:

                https://devtools.herroom.com/po_process_api.asp:
                    CALL AspTest()
                    ---------------------------------------
                    Public Sub AspTest()
                        Dim Shell
                        Set Shell = Server.CreateObject("WScript.Shell")
                        Shell.Run "e:\www\HerTools\netbin\MagentoProductAPI.exe ASPTEST"		
                    End Sub

                -- THIS PROCESS:
                    1) EXE will produce the json and push it to DB for archiving
                    2) EXE will create json file
                    3) EXE will push it to SFTP folder or Automate task can TBD

            */

            DataSet ds;
            System.Data.DataTable dtSku;
            System.Data.DataTable dtMfr;
            System.Data.DataTable dtCost;
            int SkuRow = 11;

            try
            {

                //CREATE XLS //////////////////////////////////////////////////////
                object misValue = System.Reflection.Missing.Value;

                Microsoft.Office.Interop.Excel.Application xlApp;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                xlApp = new Microsoft.Office.Interop.Excel.Application();

                //POPULATE XLS FIELDS
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                ///////////////////////////////////////////////////////////////

                OutFilename = ConfigurationManager.AppSettings["SPS850FILEOUTPUT"] + "PO" + PONumber + ".json";
                FilenameXls = ConfigurationManager.AppSettings["SPS850FILEOUTPUT"] + "PO" + PONumber + ".xlsx";

                //Datatables: 1-Sku info; 2-Mfr info
                ds = Helper.Sql_Misc_Fetch_Dataset("EXEC HerRoom..[proc_po_process_json] @PONumber = '" + PONumber + "', @MFRCode = '" + MFRCode + "'");

                // BBHer.BuyerName [BuyerHer]   , BBHis.BuyerName[BuyerHis] !!!!!!!!!!!!
                // !! NEED TO call  herroom..[proc_po_process_peachtree_report] @Date varchar(32) = '']
                // PO_PeachTree_Report() WORKS 


                if (MagetnoProductAPI.DevMode > 1)
                {
                    Console.WriteLine("DS: " + ds.Tables.Count.ToString());
                }

                if (ds.Tables.Count >= 3)
                {
                    dtSku = ds.Tables[0];
                    dtMfr = ds.Tables[1];
                    dtCost = ds.Tables[2];

                    xlWorkSheet.Name = "PO " + PONumber;

                    //XLS HEADER INFO  //////////////////////////////////////////////////////////
                    // Purchase Order 216208POTerms: 8 % 30, Net 31Day
                    // Issue Date  6 / 17 / 2024
                    // Ship Start Date 6 / 17 / 2024   Update Start Date
                    // Ship Cancel Date    7 / 2 / 2024    Update Cancel Date
                    // First Receipt   6 / 26 / 2024
                    // Last Receipt    6 / 26 / 2024
                    //
                    // AutoCancel Enabled

                    WriteXLSLine(xlWorkSheet, "Purchase Order," + ",POTerms:," + dtMfr.Rows[0]["poterms"].ToString() + ",Buyer:," + dtMfr.Rows[0]["buyername"].ToString(), 1, true);
                    WriteXLSLine(xlWorkSheet, "Issue Date," + dtMfr.Rows[0]["startdate"].ToString(), 2);
                    WriteXLSLine(xlWorkSheet, "Ship Start Date," + dtMfr.Rows[0]["postartshipdate1"].ToString(), 3);
                    WriteXLSLine(xlWorkSheet, "Ship Cancel Date," + dtMfr.Rows[0]["pocanceldate1"].ToString(), 4);
                    WriteXLSLine(xlWorkSheet, "First Receipt,-", 5);
                    WriteXLSLine(xlWorkSheet, "Last Receipt,-", 6);
                    WriteXLSLine(xlWorkSheet, " ", 7);
                    WriteXLSLine(xlWorkSheet, "AutoCancel Enabled," + dtMfr.Rows[0]["poautocancel"].ToString(), 8, true);


                    //////////////////////////////////////////////////////////

                    //Main Output
                    //OutputLine = "UPC,QtyOrdered, UnitCost,Total,Stylenumber,Desc,Color,Size";

                    OutputLine = "Orig. Line,Manufacturer,P.O.,Posted,Style,Description,Color,Color Code,size,UPC,Qty Ordered,Cost,Ext Cost,Receive,Color Override,Style Override,Closeout,SKU-Closeout,Backorder";

                    WriteXLSLine(xlWorkSheet, OutputLine, 10, true);

                    HelperModels.POEDI.PO htPO = new HelperModels.POEDI.PO();
                    HelperModels.POEDI.Header htHeader = new HelperModels.POEDI.Header();
                    HelperModels.POEDI.OrderHeader htOrderheader = new HelperModels.POEDI.OrderHeader();
                    htPO.header = new HelperModels.POEDI.Header();
                    htPO.header.OrderHeader = new HelperModels.POEDI.OrderHeader();

                    htPO.header.OrderHeader.PurchaseOrderNumber = dtMfr.Rows[0]["ponumber"].ToString();
                    htPO.header.OrderHeader.TsetPurposeCode = "00";
                    htPO.header.OrderHeader.PrimaryPOTypeCode = "SA";
                    htPO.header.OrderHeader.PurchaseOrderDate = DateTime.Parse(dtMfr.Rows[0]["startdate"].ToString());
                    //2024-07-25
                    //htPO.header.OrderHeader.Vendor = dtMfr.Rows[0]["manufacturername"].ToString();
                    htPO.header.OrderHeader.Vendor = dtMfr.Rows[0]["manufacturercode"].ToString();

                    if (IncludeOptionalCode)
                    {
                        htPO.header.OrderHeader.CustomerOrderNumber = "Sourcing will be added in the future";
                    }

                    htPO.header.Date = new List<HelperModels.POEDI.DateInfo>();
                    HelperModels.POEDI.DateInfo dateinfo010 = new HelperModels.POEDI.DateInfo();
                    dateinfo010.Datetimequalifier = "010";
                    dateinfo010.Date = DateTime.Parse(dtMfr.Rows[0]["POStartShipDate1"].ToString());
                    htPO.header.Date.Add(dateinfo010);

                    HelperModels.POEDI.DateInfo dateinfo001 = new HelperModels.POEDI.DateInfo();
                    dateinfo001.Datetimequalifier = "001";
                    dateinfo001.Date = DateTime.Parse(dtMfr.Rows[0]["POCancelDate1"].ToString());
                    htPO.header.Date.Add(dateinfo001);


                    htPO.header.address = new List<HelperModels.POEDI.Address>();
                    Addr = new HelperModels.POEDI.Address();
                    Addr.AddressTypeCode = "ST";
                    Addr.AddressLocationNumber = dtMfr.Rows[0]["POAddressLocationNumber"].ToString();     //"001";
                    Addr.LocationCodeQualifier = "92;";
                    Addr.Address1 = "8941 Empress Row #149";
                    Addr.City = "Dallas";
                    Addr.State = "TX";
                    Addr.PostalCode = "75247";

                    if (IncludeOptionalCode)
                    {
                        Addr.AddressName = "Andragroup Warehouse";
                        Addr.Country = "USA";
                    }
                    htPO.header.address.Add(Addr);

                    if (IncludeOptionalCode)
                    {
                        htPO.header.notes = new List<HelperModels.POEDI.Notes>();
                        note = new HelperModels.POEDI.Notes();
                        note.NoteCode = "GEN";
                        note.Note = "";
                        htPO.header.notes.Add(note);
                    }

                    htPO.header.Terms = new List<HelperModels.POEDI.Terms>();
                    term = new HelperModels.POEDI.Terms();
                    term.termsDescription = dtMfr.Rows[0]["POTerms"].ToString();
                    htPO.header.Terms.Add(term);

                    if (IncludeOptionalCode)
                    {
                        htPO.header.FOBRelatedInstruction = new List<HelperModels.POEDI.FOBRelatedInstruction>();
                        FOBRelatedInstr = new HelperModels.POEDI.FOBRelatedInstruction();
                        FOBRelatedInstr.FOBPayCode = "CC";
                        FOBRelatedInstr.FOBLocationQualifier = "WH";
                        FOBRelatedInstr.FOBLocationDescription = "";
                        FOBRelatedInstr.FOBTitlePassageCode = "WH";
                        FOBRelatedInstr.TransportationTermsType = "02";
                        FOBRelatedInstr.RiskOfLossCode = "IR";
                        FOBRelatedInstr.Description = "";
                        htPO.header.FOBRelatedInstruction.Add(FOBRelatedInstr);
                    }

                    //////////////////////////////////////////////////
                    if (IncludeOptionalCode)
                    {
                        htPO.ordersummary = new HelperModels.POEDI.OrderSummary();
                        htPO.ordersummary.TotalAmount = dtCost.Rows[0]["cost"].ToString();
                        htPO.ordersummary.TotalLineItemNumber = dtCost.Rows[0]["rows"].ToString();
                        htPO.ordersummary.TotalQuantity = dtCost.Rows[0]["quantity"].ToString();
                    }

                    /////////////////////////////////////////////////
                    List<HelperModels.POEDI.Item> htItems = new List<HelperModels.POEDI.Item>();
                    //HelperModels.POEDI.Item htItem;

                    List<HelperModels.POEDI.LineItem> htLineItems = new List<HelperModels.POEDI.LineItem>();
                    HelperModels.POEDI.LineItem htLineItem;

                    LIs = new HelperModels.POEDI.LineItem[dtSku.Rows.Count];

                    /////////////////////////////////////////////
                    foreach (DataRow dr in dtSku.Rows)
                    {
                        htLineItem = new HelperModels.POEDI.LineItem();
                        htLineItem.orderline = new HelperModels.POEDI.Orderline();

                        htLineItem.orderline.LineSequenceNumber = SkuRow.ToString();

                        htLineItem.orderline.VendorPartNumber = dr["upc"].ToString();
                        htLineItem.orderline.ConsumerPackageCode = dr["upc"].ToString();
                        htLineItem.orderline.OrderQty = dr["POorder"].ToString();
                        htLineItem.orderline.OrderQtyUOM = "EA";
                        htLineItem.orderline.PurchasePrice = dr["ourcost"].ToString();
                        htLineItem.orderline.Color = dr["colorname"].ToString();
                        htLineItem.orderline.Size = dr["size"].ToString();

                        if (IncludeOptionalCode)
                        {
                            htLineItem.orderline.ExtendedItemTotal = dr["extendedcost"].ToString();
                        }

                        //////////////////////////////////////////////////
                        htLineItem.productorItemDescriptions = new List<HelperModels.POEDI.ProductorItemDescriptions>();
                        ProductorItemDescription = new HelperModels.POEDI.ProductorItemDescriptions();
                        ProductorItemDescription.ProductCharacteristicCode = "08";  // dr["postyle"].ToString();
                        ProductorItemDescription.ProductDescription = dr["productname"].ToString();
                        htLineItem.productorItemDescriptions.Add(ProductorItemDescription);

                        htLineItems.Add(htLineItem);
                        SkuRow++;

                        //BUYERS XLS FILE
                        OutputLine = (SkuRow - 11).ToString() + "," + dtMfr.Rows[0]["manufacturername"].ToString() + "," + PONumber;

                        OutputLine += "," + dtMfr.Rows[0]["startdate"].ToString() + "," + dr["stylenumber"].ToString() + "," + dr["productname"].ToString();

                        OutputLine += "," + dr["colorname"].ToString() + "," + dr["colorcode"].ToString() + "," + dr["size"].ToString() + ",'" + dr["upc"].ToString();

                        OutputLine += "," + dr["poorder"].ToString() + ",$" + dr["ourcost"].ToString() + ",$" + dr["extendedcost"].ToString();

                        OutputLine += ",in process," + dr["pocoloroverride"].ToString() + "," + dr["postyleoverride"].ToString();

                        OutputLine += "," + dr["closeout"].ToString() + "," + dr["upccloseout"].ToString() + "," + dr["bo"].ToString();

                        WriteXLSLine(xlWorkSheet, OutputLine, SkuRow - 1);

                    }

                    // XLS FOOTER INFO /////////////////////////////////////////////////////////
                    WriteXLSLine(xlWorkSheet, ",Lines,Pieces,Value", SkuRow + 1, true);
                    WriteXLSLine(xlWorkSheet, "Original," + dtCost.Rows[0]["lines"].ToString() + "," + dtCost.Rows[0]["quantity"].ToString() + "," + dtCost.Rows[0]["cost"].ToString(), SkuRow + 2, true);
                    WriteXLSLine(xlWorkSheet, "Current,0,0,0", SkuRow + 3, true);
                    WriteXLSLine(xlWorkSheet, "Special Instructions," + dtMfr.Rows[0]["ponote"].ToString(), SkuRow + 4, true);
                    WriteXLSLine(xlWorkSheet, " ", SkuRow + 5);

                    htPO.LineItem = htLineItems;

                    HelperModels.POEDI.POOrder order = new HelperModels.POEDI.POOrder();
                    order.order = htPO;

                    var TM_JsonOrder = JsonConvert.SerializeObject(order, Formatting.Indented, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });

                    //Necessary for now, will work out format kinks once it is 100% vetted by SPS
                    //TM_JsonOrder = TM_JsonOrder.Replace(@"""placeholder"": ""xxx""", LIMs); 

                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine(TM_JsonOrder);
                    }

                    SaveJsonFile(TM_JsonOrder, OutFilename);

                    // INSERT Hercust..PurchaseOrder Line if all went well
                   
                }
                else
                {
                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("NO DB DATA FOUND FOR PO " + PONumber);
                    }

                    WriteXLSLine(xlWorkSheet, "NO DB DATA FOUND FOR PO " + PONumber, 1);

                    ReturnValue = false;
                }

                //FINISH EXCEL OVERHEAD /////////////////////////////////////////
                xlWorkBook.SaveCopyAs(FilenameXls);
                Sheets xlSheets = null;
                xlSheets = xlWorkBook.Sheets as Sheets;

                xlWorkBook.SaveCopyAs(FilenameXls);

                xlWorkBook.Close(false, FilenameXls, misValue);
                xlApp.Quit();
                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
            }
            catch(Exception ex)
            {
                if (MagetnoProductAPI.DevMode > 0 || MagetnoProductAPI.DevMode == -1)
                {
                    Console.WriteLine("PO XLS ERROR: " + ex.ToString());
                }
                ReturnValue = false;
            }
            /////////////////////////////////////////////////////////////////////// 

            return ReturnValue;
        }

        public static Boolean PO_PeachTree_Report()
        {
            Boolean ReturnValue = true;
            System.Data.DataTable dt;
            string OutFileName;
            //string lineout;

            try
            {

                // OutFileName = ConfigurationManager.AppSettings["PEACHTREEPATH"] + "peachtree_export_TEST.csv";
                OutFileName = ConfigurationManager.AppSettings["PEACHTREEPATH"];

                dt = Helper.Sql_Misc_Fetch("EXEC herroom..[proc_po_process_peachtree_report] @OutputFormat=2");

                if (dt.Rows.Count > 0)
                {
                    using (var tw = new StreamWriter(OutFileName, true))
                    {
                        foreach (DataRow dr in dt.Rows)
                        {
                            if (MagetnoProductAPI.DevMode > 0)
                            {
                                Console.WriteLine(dr["reportrow"].ToString());
                            }
                            tw.WriteLine(dr["reportrow"].ToString());
                        }

                        tw.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine("ERROR: Peachtree CSV " + ex.ToString());
                }
                ReturnValue = false;
            }

            return ReturnValue;
        }

        public static Boolean PO_PDFLabels_Send(string MfrCode, string PONumber) //, string EmailAddresses)
        {
            Boolean ReturnValue = true;
            DirectoryInfo di;
            System.Data.DataTable dt;
            string EmailBody= "";

            dt = Helper.Sql_Misc_Fetch("SELECT REPLACE(ISNULL(POemail,''),',',';') [POemail], ISNULL(POemailname,'') [POemailname], ISNULL(SigFileName,'') [SigFileName], ISNULL(Email,'') [ClerkEmail], ClerkName FROM Herroom..Manufacturers RMM LEFT OUTER JOIN Herroom..PoClerks POC ON POC.id = RMM.PoClerkID  WHERE Manufacturercode = '" + MfrCode + "' AND SendLabels = 1");

            //find last file in \\10.10.1.113\e$\www\barcodeGen3\PDF_Files\Shdw01, email is - 

            //DateTime LastFile = Directory.GetLastWriteTime(@"\\10.10.1.113\e$\www\barcodeGen3\PDF_Files\" + MfrCode);
            DateTime LastFile = Directory.GetLastWriteTime(ConfigurationManager.AppSettings["SPS850UPCLABELSFROM"] + MfrCode);

            //di = new DirectoryInfo(@"\\10.10.1.113\e$\www\barcodeGen3\PDF_Files\" + MfrCode);
            di = new DirectoryInfo(ConfigurationManager.AppSettings["SPS850UPCLABELSFROM"] + MfrCode);
            FileInfo[] Files = di.GetFiles();

            if (MagetnoProductAPI.DevMode > 1)
            {
                Console.WriteLine(LastFile.ToLongDateString() + " " + LastFile.ToLongTimeString());
            }

            //Move Files to be Processed to Root Directory
            foreach (FileInfo fi in Files)
            {
                if (fi.LastWriteTime.ToLongDateString() + " " + fi.LastWriteTime.ToLongTimeString() == LastFile.ToLongDateString() + " " + LastFile.ToLongTimeString())
                {
                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("COPY " + fi.Name + "; " + fi.LastWriteTime.ToLongDateString() + "; " + LastFile.ToLongDateString());
                    }
                    

                    if (dt.Rows.Count > 0 && dt.Rows[0]["POemail"].ToString().Length > 0)
                    {
                        EmailBody = "Purchase Order for HerRoom/HisRoom PO#: " + PONumber + " <br /><br />"
                            + "Authorized By:<br><img src=''http://tools.herroom.com/images/" + dt.Rows[0]["SigFileName"].ToString() + "'' border=0 width=244 height=91>''<br /><br />"
                            + dt.Rows[0]["ClerkName"].ToString() + ", Purchasing Agent<br><a href=''mailto:" + dt.Rows[0]["ClerkEmail"].ToString() + "</a>''<br/>214-691-1191 <br /><br /> " 
                            + "Please see attachment<br />";

                        //email it
                        if (MagetnoProductAPI.DevMode > 1)
                        {
                            Console.WriteLine("EmailBody: " + EmailBody);
                            Console.WriteLine("---- fi.FullName: " + fi.FullName);

                            Console.WriteLine("INSERT communications..PO_EMAIL_Attachments(Recipients, CC, BCC, EmailBody , EmailSubject, Files, Datestamp , EmailSent) "
                                + " SELECT '" + dt.Rows[0]["poemail"].ToString() + "', 'buyers@andragroup.com', 'thomas@andragroup.com', '" + EmailBody + "', 'PO Shipping Labels for PO:" + PONumber + "', '" + fi.FullName + "', Getdate(), 0");
                        }
                        else
                        {
                            // DOES NOT WORK, Proxy SQL Email file attachment issue, don't want to redo SQL security at this time
                            // Helper.Sql_Misc_NonQuery("EXEC communications..[proc_mag_email_po_shippinglabel_send] @recipients = '" + dt.Rows[0]["poemail"].ToString() + "' "
                            //Helper.Sql_Misc_NonQuery("EXEC communications..[proc_mag_email_po_shippinglabel_send] @recipients = 'thomas@andragroup.com;thomas.tribble@gmail.com;' "
                            //    + " , @cc_recipients  = 'buyers@andragroup.com' "
                            //    + " , @subject = 'PO Shipping Labels for PO:" + PONumber + "' "
                            //    + " , @body = '" + EmailBody + "'"
                            //    + " , @file_attachments = '" + fi.FullName + "'");

                            //Hourly SQL Job runs to send these with attachements
                            Helper.Sql_Misc_NonQuery("INSERT communications..PO_EMAIL_Attachments(Recipients, CC, BCC, EmailBody , EmailSubject, Files, Datestamp , EmailSent) "
                                + " SELECT '" + dt.Rows[0]["poemail"].ToString() + "', 'buyers@andragroup.com', 'thomas@andragroup.com', '" + EmailBody + "', 'PO Shipping Labels for PO:" + PONumber + "', '" + fi.FullName + "', Getdate(), 0");

                       }


                    }
                    else
                    {
                        //EMAIL BUYERS for failure ??
                    }
                }
                //Email if newest file - don't do for now 
                //System.IO.File.Move(@"\\10.10.1.113\e$\www\barcodeGen3\PDF_Files\" + MfrCode + @"\" + fi.Name, @"\\10.10.1.113\e$\www\barcodeGen3\PDF_Files\" + MfrCode + @"\archive\" + fi.Name);
                //fi.MoveTo(@"\\10.10.1.113\e$\www\barcodeGen3\PDF_Files\" + MfrCode + @"\archive\" + fi.Name);
            }

            return ReturnValue;
        }
   
        public static Boolean Incoming_856File_Process()
        {
            Boolean ReturnValue = true;
            DirectoryInfo di;
            DirectoryInfo diProcess;
            string FileFolder;
            String DirectoryArchive;
            String DirectoryError;
            String DirectoryFTP;
            System.Data.DataTable dt;

            //find all files, process the unprocessed ones;
            FileFolder = ConfigurationManager.AppSettings["SPS856FILEINCOMING"];
            DirectoryArchive = ConfigurationManager.AppSettings["SPS856FILEINCOMINGARCHIVE"];
            DirectoryError = ConfigurationManager.AppSettings["SPS856FILEINCOMINGERROR"];
            DirectoryFTP = ConfigurationManager.AppSettings["SPS856FILEINCOMINGFTP"];

            di = new DirectoryInfo(DirectoryFTP);
            FileInfo[] Files = di.GetFiles();

            string POnumber;

            //Move Files to be Processed to Root Directory
            foreach (FileInfo fi in Files)
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine("File: " + fi.FullName + "; " + fi.FullName.Replace(@"\FTP", ""));
                    Console.WriteLine(" ------------- "); 
                }

                dt = new System.Data.DataTable();  // clear table 
                dt = Helper.Sql_Misc_Fetch("SELECT COUNT(*) [cnt] FROM hercust..po_856_incoming_file WHERE Filename = '" + fi.FullName.Replace(@"\FTP", "") + "'");
                if (dt.Rows.Count > 0 && dt.Rows[0]["cnt"].ToString() == "0")
                {
                    //Move File to Main folder to process
                    fi.MoveTo(FileFolder + fi.Name);
                }
            }

            diProcess = new DirectoryInfo(FileFolder);
            FileInfo[] FilesProcess = diProcess.GetFiles();

            foreach (FileInfo fi in FilesProcess)
            {
                if (fi.Name.EndsWith(".xlsx"))
                {
                    continue; 
                }

                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine(" !! PROCESS File: " + fi.FullName + "; " + fi.Name);
                }

                Helper.Sql_Misc_NonQuery("INSERT hercust..po_856_incoming_file([Filename], processed, datestamp) SELECT '" + fi.FullName + "', 100, Getdate()  WHERE 0 = (SELECT COUNT(*) FROM hercust..po_856_incoming_file WHERE [Filename] = '" + fi.FullName + "') ");              

                //if (Incoming_856File(fi.FullName, fi.Name))
               if (POProcess.Incoming_856File(fi.FullName, fi.Name, out POnumber))
                {
                    fi.MoveTo(DirectoryArchive + fi.Name);
                }
                else
                {
                    fi.MoveTo(DirectoryError + fi.Name);
                }
            }

            return ReturnValue;
        }

        public static Boolean Incoming_856File(string FilePath, string ShortFileName)
        {
            Boolean ReturnValue = true;
            String jsontext;
            String Lineout = "";
            int Rowcount = 2;
            System.Data.DataTable dtPOA;
            System.Data.DataTable dt856Incoming;
            decimal CartonsinPO;
            string FilenameXls;
            string Address_Address1 = "unknown";
            string Address_City = "unknown";
            string Address_Postcode = "unknown";
            string Address_Name = "unknown";
            string Carrier = "";
            string Redded;
            string FileArchive;

            //FileArchive = ConfigurationManager.AppSettings["SPS856FILEARCHIVE"];
            FileArchive = FilePath.Replace(@"/incoming/", @"/incoming/archive/");

            //Set to status 200 while processing, then 600, 700
            Helper.Sql_Misc_NonQuery("UPDATE HerCust..[PO_856_Incoming_file] SET Processed = 200 WHERE Filename = '" + ShortFileName + "';");

            FilenameXls = FilePath.Replace(".json", ".xlsx").Replace(".txt", ".xlsx");
            StreamReader sr = new StreamReader(FilePath);
            jsontext = sr.ReadToEnd();
            sr.Close();

            HelperModels.SPS856.SHFile.File856 SP = new HelperModels.SPS856.SHFile.File856();

            SP = JsonConvert.DeserializeObject<HelperModels.SPS856.SHFile.File856>(jsontext, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
        
            //FIND ShipFrom Address Info
            foreach (HelperModels.SPS856.Address AH in SP.Header.Address)
            {
                if (AH.AddressTypeCode.ToUpper() == "SF")
                {
                    Address_Address1 = AH.Address1;
                    Address_City = AH.City;
                    Address_Postcode = AH.PostalCode;
                    Address_Name = AH.AddressName;
                }
            }

            if (SP.Header.CarrierInformation.Count > 0)
            {
                Carrier = SP.Header.CarrierInformation[0].CarrierAlphaCode;
            }

            //CREATE XLS //////////////////////////////////////////////////////////
            object misValue = System.Reflection.Missing.Value;

            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            xlApp = new Microsoft.Office.Interop.Excel.Application();

            //POPULATE XLS FIELDS
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            /////////////////////////////////////////////////////////////////////////

            //Lineout = "TransactionType,Buyer,PONum,PODate,OurCost,QuantityShipped, QuantityOrdered, ImportedDate, AccountingId, TrackingNumber,  ScheduledDelivery, ShipDate, ShipVia,  NumOfCartonsShipped, ShipFromName, ShipFromAddressLineOne, ShipFromAddressLineTwo, ShipFromCity, ShipFromState,ShipFromZip,ShipFromCountry,ShipFromAddressCode,VendorNum,DCCode,TransportationMethod,Status,UOMOfUPCs,Item Description,UPC,POcost,ManufacturerName,StyleNumber,BO";

            Lineout = "TransactionType,PONum,Buyer,ManufacturerName,PODate,PO Start Ship Date, PO Cancel Date,Stylenumber,Item Description,Color,Color Code,Size,UPC,OurCost,QuantityShipped,QuantityOrdered,BO,ImportedDate,AccountingId,TrackingNumber,ScheduledDelivery,Vendor ShipDate,ShipVia,NumOfCartonsShipped,ShipFromName,ShipFromAddressLineOne, ShipFromAddressLineTwo,ShipFromCity,ShipFromState,ShipFromZip,ShipFromCountry,ShipFromAddressCode,VendorNum,DCCode,TransportationMethod,Status,UOMOfUPCs,POcost,LineNum,Notes";

            if (MagetnoProductAPI.DevMode > 0)
            {
                Console.WriteLine(Lineout);
            }
            WriteXLSLine(xlWorkSheet, Lineout, 1);

            foreach (HelperModels.SPS856.OrderLevel OL in SP.OrderLevel)
            {
                try
                {
                    // !! WILL NEED PAGE FOR EACH PO ??? !!!
                    xlWorkSheet.Name = "PO " + OL.OrderHeader.PurchaseOrderNumber;

                    Lineout = "";
                    foreach (HelperModels.SPS856.PackLevel PL in OL.PackLevel)
                    {
                        CartonsinPO = 0;
                        foreach (HelperModels.SPS856.ItemLevel ILTotal in PL.ItemLevel)
                        {
                            CartonsinPO += ILTotal.ShipmentLine.ShipQty;
                        }

                        foreach (HelperModels.SPS856.ItemLevel IL in PL.ItemLevel)
                        {
                            try
                            {
                                dtPOA = Helper.Sql_Misc_Fetch("SELECT PA.*,  SS.ManufacturerCode, (UnitCost * QtyRcvd) [pocost]   "
                                    + ", 0[QtyOfUPCs within Pack], RMM.manufacturerName "
                                    + ", BB.buyername, SS.ProductName + ' ' + II.size + '-' + II.ColorCode[productname] "
                                    + ", CASE WHEN(SELECT COUNT(*) FROM Hercust..Orders OO INNER JOIN Hercust..Items CII ON ordernum = orderno "
                                    + "         WHERE SKU = II.UPC AND CII.Backorder = 1) > 0 THEN 'x' ELSE '' END [bo], II.stylenumber, II.ourcost, PO.StartDate, PO.CancelDate, II.ColorCode, ColorName, II.size "
                                    + " , (SELECT TOP 1 LineNum FROM HerCust..Expected EX WHERE EX.PONumber = PA.PONumber AND Ex.upc = PA.UPC ORDER BY EX.Invkey DESC) [LineNum] "
                                    + " FROM Hercust..PoArchive PA with (nolock) "
                                    + " LEFT OUTER JOIN Hercust..PurchaseOrders PO with (nolock) ON PO.PoNum = PA.PONumber"
                                    + " LEFT OUTER JOIN Herroom..Items II with (nolock) ON II.upc = PA.UPC "
                                    + " LEFT OUTER JOIN Herroom..Styles SS with (nolock) ON SS.StyleNumber = II.StyleNumber "
                                    + " LEFT OUTER JOIN Herroom..Manufacturers RMM with (nolock) ON RMM.ManufacturerCode = SS.ManufacturerCode "
                                    + " LEFT OUTER JOIN Herroom..Buyers BB with (nolock) ON BB.id = CASE WHEN SS.Gender = 'm' THEN RMM.BuyerHisID ELSE RMM.BuyerID END "
                                    + " LEFT OUTER JOIN Herroom..Colors CO with (nolock) ON CO.ColorCode = II.ColorCode "
                                    + " WHERE PA.PONumber = '" + OL.OrderHeader.PurchaseOrderNumber + "' AND PA.UPC = '" + IL.ShipmentLine.ConsumerPackageCode + "'");

                                //+ "WHERE PA.PONumber = '212249' AND PA.UPC = '608926187876'");

                                //Get exising UPC data for PO
                                dt856Incoming = Helper.Sql_Misc_Fetch("SELECT * FROM Hercust..po_856_incoming WHERE POnumber = '" + OL.OrderHeader.PurchaseOrderNumber + "' AND UPC = '" + IL.ShipmentLine.ConsumerPackageCode + "' AND ShippingSerialID <> '" + OL.PackLevel[0].Pack.ShippingSerialID + "'; ");

                                if (dtPOA.Rows.Count > 0)
                                {
                                    DataRow drPOA = dtPOA.Rows[0];

                                    //Lineout = "856," + OL.OrderHeader.Vendor + "," + PL.Pack.ShippingSerialID + ",," + IL.ShipmentLine.ConsumerPackageCode + ",," + SP.Header.ShipmentHeader.ShipDate + "," +SP.Header.ShipmentHeader.CurrentScheduledDeliveryDate + ",ANDRAGROUP,DBA:HERROOM.COM,8941 EMPRESS ROW,DALLAS,TX,75247,,01,,Assigned by Buyer or Buyer's Agent,CTN25,-Cost-,,3,,," + OL.OrderHeader.Vendor + ",,,,,,,,,01,,,,,-PONumber-,-Date??-,,--Cost--,ANDRA GROUP,01,,,1,CTN25,--000074919--,,,,,,-cost-,--line#--,Each,ANDRA GROUP,01," + IL.ShipmentLine.LineSequenceNumber + ",,SKU,," + IL.ShipmentLine.ShipQty.ToString() + ",Each," + IL.ShipmentLine.ShipQty.ToString() + ",,,Each";
                                    //TransactionType	PONum	PODate	InvoiceNum	OurCost	QuantityShipped	QuantityOrdered	ImportedDate	AccountingId	TrackingNumber	BillOfLading	ScheduledDelivery	ShipDate	ShipVia	ShipToType	PackagingType	GrossWeight	GrossWeightUOM	NumOfCartonsShipped	ShipFromName	ShipFromAddressLineOne	ShipFromAddressLineTwo	ShipFromCity	ShipFromState	ShipFromZip	ShipFromCountry	ShipFromAddressCode	VendorNum	DCCode	TransportationMethod	Status	TimeShipped	OrderWeight	QtyOfUPCs within Pack	UOMOfUPCs	Item Description	UPC	POcost	ManufacturerName	Buyer	StyleNumber	BO

                                    Lineout = "856," + OL.OrderHeader.PurchaseOrderNumber + "," + drPOA["buyername"].ToString() + "," + drPOA["ManufacturerName"].ToString() + "," + drPOA["startdate"].ToString() + "," + drPOA["canceldate"].ToString() + "," + OL.OrderHeader.PurchaseOrderDate + "," + drPOA["stylenumber"].ToString() + "," + drPOA["productname"].ToString() + "," + drPOA["colorname"].ToString() + ", " + drPOA["colorcode"].ToString() + "," + drPOA["size"].ToString() + ",'" + drPOA["upc"].ToString();

                                    Lineout += ",$" + drPOA["ourcost"].ToString() + "," + IL.ShipmentLine.ShipQty;

                                    //Lineout += "," + drPOA["NumberExp"].ToString() + "," + drPOA["bo"].ToString() + "," + drPOA["datein"].ToString() + "," + drPOA["ManufacturerCode"].ToString() + ",'" + OL.PackLevel[0].Pack.ShippingSerialID + ", " + SP.Header.ShipmentHeader.CurrentScheduledDeliveryDate + ", " + SP.Header.ShipmentHeader.ShipDate + ",";
                                    Lineout += "," + drPOA["NumberExp"].ToString() + "," + drPOA["bo"].ToString() + "," + drPOA["datein"].ToString() + "," + drPOA["ManufacturerCode"].ToString() + ",'" + SP.Header.ShipmentHeader.BillOfLadingNumber  + ", " + SP.Header.ShipmentHeader.CurrentScheduledDeliveryDate + ", " + SP.Header.ShipmentHeader.ShipDate + ",";

                                    Lineout += "," + CartonsinPO + "," + Address_Name;   //OL.OrderHeader.Vendor;

                                    Lineout += "," + Address_Address1 + ",," + Address_City + ",," + Address_Postcode + ",,,,01," + Carrier + ",,Each,$" + drPOA["pocost"].ToString();

                                    Lineout += "," + drPOA["LineNum"].ToString();       // LineNum (before Notes)

                                    if (MagetnoProductAPI.DevMode > 0)
                                    {
                                        Console.WriteLine(Lineout);
                                    }

                                    Redded = "";

                                    if (IL.ShipmentLine.ShipQty != decimal.Parse(drPOA["NumberExp"].ToString()))
                                    {
                                        Redded = "O,";
                                    }
                                    if (drPOA["LineNum"].ToString() != IL.ShipmentLine.LineSequenceNumber.ToString())
                                    {
                                        Redded += "AM,";
                                    }

                                    //if (IL.ShipmentLine.ShipQty != decimal.Parse(drPOA["NumberExp"].ToString()))
                                    //{
                                    //    WriteXLSLine(xlWorkSheet, Lineout, Rowcount, false, "O");
                                    //}
                                    //if (Redded.Length > 0)
                                    //{
                                    WriteXLSLine(xlWorkSheet, Lineout, Rowcount, false, Redded);
                                    //}
                                    //else
                                    //{
                                    //   WriteXLSLine(xlWorkSheet, Lineout, Rowcount);
                                    //}

                                    Rowcount++;

                                    if (MagetnoProductAPI.DevMode > 0)
                                    {
                                        Console.WriteLine("INSERT Hercust..PO_856_Incoming(PONumber, FileName, ShippingSerialID, UPC, QtyIn, Notes, Datestamp, BillOfLadingNumber) SELECT '" + OL.OrderHeader.PurchaseOrderNumber + "', '" + ShortFileName + "', '" + PL.Pack.ShippingSerialID + "', '" + IL.ShipmentLine.ConsumerPackageCode + "', " + IL.ShipmentLine.ShipQty.ToString() + ", '', '" + SP.Header.ShipmentHeader.ShipDate + "' WHERE 0 = (SELECT COUNT(*) FROM Hercust..PO_856_Incoming WHERE PONumber = '" + OL.OrderHeader.PurchaseOrderNumber + "' AND ShippingSerialId = '" + PL.Pack.ShippingSerialID + "' AND UPC = '" + IL.ShipmentLine.ConsumerPackageCode  + "', '" + SP.Header.ShipmentHeader.BillOfLadingNumber + "')");
                                        Console.WriteLine("------");
                                    }

                                    //Helper.Sql_Misc_NonQuery("INSERT Hercust..PO_856_Incoming(PONumber, FileName, ShippingSerialID, UPC, QtyIn, Notes, Datestamp, BillOfLadingNumber) SELECT '" + OL.OrderHeader.PurchaseOrderNumber + "', '" + ShortFileName + "', '" + PL.Pack.ShippingSerialID + "', '" + IL.ShipmentLine.ConsumerPackageCode + "', " + IL.ShipmentLine.ShipQty.ToString() + ", '', '" + SP.Header.ShipmentHeader.ShipDate + "', '" + SP.Header.ShipmentHeader.BillOfLadingNumber + "' WHERE 0 = (SELECT COUNT(*) FROM Hercust..PO_856_Incoming WHERE PONumber = '" + OL.OrderHeader.PurchaseOrderNumber + "' AND ShippingSerialId = '" + PL.Pack.ShippingSerialID + "' AND UPC = '" + IL.ShipmentLine.ConsumerPackageCode + "')");
                                    Helper.Sql_Misc_NonQuery("INSERT Hercust..PO_856_Incoming(PONumber, FileName, ShippingSerialID, UPC, QtyIn, Notes, Datestamp, BillOfLadingNumber) SELECT '" + OL.OrderHeader.PurchaseOrderNumber + "', '" + ShortFileName + "', '" + PL.Pack.ShippingSerialID + "', '" + IL.ShipmentLine.ConsumerPackageCode + "', " + IL.ShipmentLine.ShipQty.ToString() + ", '', '" + SP.Header.ShipmentHeader.ShipDate + "', '" + SP.Header.ShipmentHeader.BillOfLadingNumber + "' WHERE 0 = (SELECT COUNT(*) FROM Hercust..PO_856_Incoming WHERE PONumber = '" + OL.OrderHeader.PurchaseOrderNumber + "' AND ShippingSerialId = '" + PL.Pack.ShippingSerialID + "' AND UPC = '" + IL.ShipmentLine.ConsumerPackageCode + "')");


                                    //Insert Existing Info for other 856s alread in DB For same PO/UPC 
                                    foreach (DataRow dr856 in dt856Incoming.Rows)
                                    {
                                        Lineout = "Previous 856," + OL.OrderHeader.PurchaseOrderNumber + "," + drPOA["buyername"].ToString() + "," + drPOA["ManufacturerName"].ToString() + "," + drPOA["startdate"].ToString() + "," + drPOA["canceldate"].ToString() + "," + OL.OrderHeader.PurchaseOrderDate + "," + drPOA["stylenumber"].ToString() + "," + drPOA["productname"].ToString() + "," + drPOA["colorname"].ToString() + ", " + drPOA["colorcode"].ToString() + "," + drPOA["size"].ToString() + ",'" + drPOA["upc"].ToString();

                                        Lineout += ",$" + drPOA["ourcost"].ToString() + "," + dr856["qtyin"].ToString();

                                        //Lineout += "," + drPOA["NumberExp"].ToString() + "," + drPOA["bo"].ToString() + "," + dr856["Datestamp"].ToString() + "," + drPOA["ManufacturerCode"].ToString() + ",'" + dr856["ShippingSerialID"].ToString() + ", " + SP.Header.ShipmentHeader.CurrentScheduledDeliveryDate + ", " + SP.Header.ShipmentHeader.ShipDate + ",";
                                        Lineout += "," + drPOA["NumberExp"].ToString() + "," + drPOA["bo"].ToString() + "," + dr856["Datestamp"].ToString() + "," + drPOA["ManufacturerCode"].ToString() + ",'" + dr856["BillOfLadingNumber"].ToString() + ", " + SP.Header.ShipmentHeader.CurrentScheduledDeliveryDate + ", " + SP.Header.ShipmentHeader.ShipDate + ",";

                                        Lineout += "," + CartonsinPO + "," + Address_Name;   //OL.OrderHeader.Vendor;

                                        Lineout += "," + Address_Address1 + ",," + Address_City + ",," + Address_Postcode + ",,,,01," + Carrier + ",,Each,$" + drPOA["pocost"].ToString();

                                        Lineout += ",?";          // LineNum ???

                                        Lineout += ",UPC FOR PO ALREADY PROCESSED FROM ANOTHER 856";

                                        if (MagetnoProductAPI.DevMode > 0)
                                        {
                                            Console.WriteLine(Lineout);
                                        }

                                        WriteXLSLine(xlWorkSheet, Lineout, Rowcount, true, "M");
                                        Rowcount++;
                                    }
                                }
                                else
                                {
                                    //Missing SKU
                                    try
                                    {
                                        Lineout = "856," + OL.OrderHeader.PurchaseOrderNumber + ",-,-,-,-," + OL.OrderHeader.PurchaseOrderDate + ",-,";

                                        Lineout += IL.ProductOrItemDescription[0].ProductDescription + ",";

                                        Lineout += IL.ShipmentLine.ProductColorDescription + ",-," + IL.ShipmentLine.ProductSizeDescription + ",'" + IL.ShipmentLine.ConsumerPackageCode;

                                        Lineout += ",-," + IL.ShipmentLine.ShipQty;

                                        //Lineout += ",0,-,-,-,'" + OL.PackLevel[0].Pack.ShippingSerialID + ", " + SP.Header.ShipmentHeader.CurrentScheduledDeliveryDate + ", " + SP.Header.ShipmentHeader.ShipDate + ",";
                                        Lineout += ",0,-,-,-,'" + SP.Header.ShipmentHeader.BillOfLadingNumber + ", " + SP.Header.ShipmentHeader.CurrentScheduledDeliveryDate + ", " + SP.Header.ShipmentHeader.ShipDate + ",";

                                        Lineout += "," + CartonsinPO + "," + Address_Name;   //OL.OrderHeader.Vendor;

                                        Lineout += "," + Address_Address1 + ",," + Address_City + ",," + Address_Postcode + ",,,,01," + Carrier + ",,Each,,,-";

                                        WriteXLSLine(xlWorkSheet, Lineout, Rowcount, true, "A");
                                        Rowcount++;

                                        if (MagetnoProductAPI.DevMode > 0)
                                        {
                                            Console.WriteLine(Lineout);
                                        }

                                        //Helper.Sql_Misc_NonQuery("INSERT Hercust..PO_856_Incoming(PONumber, FileName, ShippingSerialID, UPC, QtyIn, Notes, Datestamp, BillOfLadingNumber, CSVLine) SELECT '" + OL.OrderHeader.PurchaseOrderNumber + "', '" + ShortFileName + "', '" + PL.Pack.ShippingSerialID + "', '" + IL.ShipmentLine.ConsumerPackageCode + "', " + IL.ShipmentLine.ShipQty.ToString() + ", 'SKU NOT FOUND IN: " + IL.ShipmentLine.ConsumerPackageCode + ": " + IL.ProductOrItemDescription[0].ProductDescription + "', '" + SP.Header.ShipmentHeader.ShipDate + "' WHERE 0 = (SELECT COUNT(*) FROM Hercust..PO_856_Incoming WHERE PONumber = '" + OL.OrderHeader.PurchaseOrderNumber + "' AND ShippingSerialId = '" + PL.Pack.ShippingSerialID + "', '" + SP.Header.ShipmentHeader.BillOfLadingNumber + "', '" + Lineout.Replace("'", "''") + "')");
                                    }
                                    catch (Exception EXFail)
                                    {
                                        if (MagetnoProductAPI.DevMode > 0)
                                        {
                                            Console.WriteLine("Missing Sku XLSx Write FAIL: " + EXFail.ToString());
                                        }
                                        ReturnValue = false;
                                    }
                                }

                                Helper.Sql_Misc_NonQuery("INSERT Hercust..PO_856_Incoming(PONumber, FileName, ShippingSerialID, UPC, QtyIn, Notes, Datestamp, BillOfLadingNumber, CSVLine) SELECT '" + OL.OrderHeader.PurchaseOrderNumber + "', '" + ShortFileName + "', '" + PL.Pack.ShippingSerialID + "', '" + IL.ShipmentLine.ConsumerPackageCode + "', " + IL.ShipmentLine.ShipQty.ToString() + ", 'SKU NOT FOUND IN: " + IL.ShipmentLine.ConsumerPackageCode + ": " + IL.ProductOrItemDescription[0].ProductDescription + "', '" + SP.Header.ShipmentHeader.ShipDate + "' WHERE 0 = (SELECT COUNT(*) FROM Hercust..PO_856_Incoming WHERE PONumber = '" + OL.OrderHeader.PurchaseOrderNumber + "' AND ShippingSerialId = '" + PL.Pack.ShippingSerialID + "', '" + SP.Header.ShipmentHeader.BillOfLadingNumber + "', '" + Lineout.Replace("'", "''") + "')");

                            }
                            catch (Exception exInnerLoop)
                            {
                                if (MagetnoProductAPI.DevMode > 0)
                                {
                                    Console.WriteLine("ERROR INNER LOOP: " + exInnerLoop.ToString());
                                }
                                ReturnValue = false;
                            }
                        }
                    }
                }
                catch (Exception exOuterLoop)
                {
                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("ERROR Outer LOOP: " + exOuterLoop.ToString());
                    }
                    ReturnValue = false;
                }
            }

            //FINISH EXCEL OVERHEAD /////////////////////////////////////////
            xlWorkBook.SaveCopyAs(FilenameXls);
            Sheets xlSheets = null;
            xlSheets = xlWorkBook.Sheets as Sheets;

            xlWorkBook.SaveCopyAs(FilenameXls);

            xlWorkBook.Close(false, FilenameXls, misValue);
            xlApp.Quit();
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
            /////////////////////////////////////////////////////////////////////// 

            if (ReturnValue)
            {
                Helper.Sql_Misc_NonQuery("UPDATE Hercust..[PO_856_Incoming_file] SET [Processed] = 600 WHERE FileName = '" + ShortFileName + "'");
                File.Move(FilePath, FilePath.ToLower().Replace(@"\incoming\", @"\incoming\archive\"));
            }
            else
            {
                Helper.Sql_Misc_NonQuery("UPDATE Hercust..[PO_856_Incoming_file] SET [Processed] = 700 WHERE FileName = '" + ShortFileName + "'");
                File.Move(FilePath, FilePath.ToLower().Replace(@"\incoming\", @"\incoming\error\"));
            }

            return ReturnValue;
        }

        public static Boolean PO_846_Process_IncomingFile(string Filename)
        {
            Boolean ReturnValue = true;
            HelperModels.SP846.File846.ItemRegistry SP = new HelperModels.SP846.File846.ItemRegistry();
            HelperModels.SP846.File846 xx = new HelperModels.SP846.File846();

            //HelperModels.SP846_File LineItemData;
            string jsontext;

            
            try
            {
                StreamReader sr = new StreamReader(Filename);
                jsontext = sr.ReadToEnd();
                sr.Close();

                if (MagetnoProductAPI.DevMode > 1)
                {
                    Console.WriteLine(jsontext);
                }

                //HelperModels.SPS856.SHFile.File856 SP
                //HelperModels.SP846_File xx = new HelperModels.SP846_File();
              

                SP = JsonConvert.DeserializeObject<HelperModels.SP846.File846.ItemRegistry>(jsontext, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
            
                if (MagetnoProductAPI.DevMode > 0)
                {
                    //Console.WriteLine(SP.Header.type);
                }

                //Write to DB
                //LineItemData = new HelperModels.SP846_File();

                //PO_846_Table_Insert(LineItemData);
            }
            catch(Exception ex)
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine("ERROR: " + ex.ToString());
                }
                ReturnValue = false;
            }

            return ReturnValue;
        }

        private static Boolean PO_846_Table_Insert(string FileName, HelperModels.SP846.File846.ItemRegistry LineItemData )
        {
            Boolean ReturnValue = true;
            string Sql = "";

            try
            {
                /*
                Sql = "EXEC HERcust..[proc_mag_po_846_incomingfile_insert]  @FileInfo = '" + FileName + "'"
                        + ", @InventoryDate = '" + LineItemData.Header.HeaderReport.InventoryDate + " " + LineItemData.Header.HeaderReport.InventoryTime + "'"
                        + ", @sku = '" + LineItemData.Structure.LineItem.?? + "'"
                        + ", @TotalQty = " + LineItemData.TotalQty.ToString()   
                        + ", @TotalQtyUOM = " + LineItemData.TotalQtyUOM.ToString()  
                        + ", @LineItemDate = '" + LineItemData.LineItemDate + "'"
                        + ", @DocumentId =  '" + LineItemData.DocumentId + "'"
                        + ", @AddressInfo = '" + LineItemData.AddressInfo + "'"
                        + ", @UnitPrice = " + LineItemData.UnitPrice  
                        + ", @LineSequenceNumber = '" + LineItemData.LineSequenceNumber + "'"
                        + ", @ProductDescription = '" + LineItemData.ProductDescription + "'"
                        + ", @VendorPartNumber = '" + LineItemData.VendorPartNumber + "'"
                        + ", @ConsumerPackageCode = '" + LineItemData.ConsumerPackageCode + "'"
                        + ", @ProductSizeDescription = '" + LineItemData.ProductSizeDescription + "'"
                        + ", @ProductColorDescription =  '" + LineItemData.ProductColorDescription + "'";
    
                if (MagetnoProductAPI.DevMode == 0)
                {
                    Helper.Sql_Misc_NonQuery(Sql);
                }

                */

            }
            catch (Exception ex)
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine("ERROR: " + ex.ToString());
                }

                ReturnValue = false;
            }

             return ReturnValue;
            }


        ///////////////////////////////////////////////////////////////////////////////
        //EXCEL Functions
        /////////////////////////////////////////////////////////////////////////////

        public static void WriteXLSLine(Microsoft.Office.Interop.Excel.Worksheet wksheet, string CSVLineout, int Rowcount, Boolean BoldText = false, String CellAlertColumn = "" )
        {
            string[] LineoutSplit;
            string[] Redded;

            LineoutSplit = CSVLineout.Split(Char.Parse(","));
            for (int xx = 0; xx < LineoutSplit.Length; xx++)
            {
                wksheet.Cells[Rowcount, xx + 1] = LineoutSplit[xx];  
            }

            if (BoldText)
            {
                wksheet.Range["A" + Rowcount.ToString() + ":S" + Rowcount.ToString()].Font.Bold = true;
            }

            
            if (CellAlertColumn != "")
            {
               /// wksheet.Range[CellAlertColumn + Rowcount.ToString() + ":" + CellAlertColumn + Rowcount.ToString()].Font.Color = System.Drawing.Color.Red;

                Redded = CellAlertColumn.Split(Char.Parse(","));
                for (int xx=0; xx < Redded.Length-1; xx++)
                {          
                    Console.WriteLine("redded [" + Redded[xx] + Rowcount.ToString() + ":" + Redded[xx] + Rowcount.ToString() + "]");
                    wksheet.Range[Redded[xx] + Rowcount.ToString() + ":" + Redded[xx] + Rowcount.ToString()].Font.Color = System.Drawing.Color.Red;
                }
                //  foreach (string rd in Redded)
                //{
                //    Console.WriteLine("redded [" + rd + Rowcount.ToString() + ":" + rd + Rowcount.ToString() + "]");
                //    wksheet.Range[rd + Rowcount.ToString() + ":" + rd + Rowcount.ToString()].Font.Color = System.Drawing.Color.Red;
                //}

            }

            //wksheet.Range[""].Font.Italic

            //if (WrapText)
            //{
            //    wksheet.Range["A1:I1"].WrapText = true;
            //}
        }

        public static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception EE)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        public static Boolean SaveJsonFile(string Json, string Filename)
        {
            Boolean Returnvalue = true;

            //if (!File.Exists(Filename))
           // {
                //writes to file
            //    System.IO.File.WriteAllText(Filename, "Text to add to the file\n");
            //}
            //else
            //{
                //File.Create(Filename);

                //using (FileStream fs = File.Create(Filename))
                //{
                //    File.WriteAllText(Filename, Json);
                //    fs.Close();
                //}

                using (var tw = new StreamWriter(Filename, true))
                {
                    tw.WriteLine(Json);
                    tw.Close();
                }

            //}

            return Returnvalue;
        }

    }
}
