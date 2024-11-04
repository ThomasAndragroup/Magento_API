using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace MagentoProductAPI
{
    class MagetnoProductAPI
    {
        public static int DevMode = 0;

        static void Main(string[] args)
        {
            long Arg2Int;

            String Arg2 = "";
            if (args.Length == 2)
            {
                Arg2 = args[1];
            }

            if (args.Length > 0)
            {
                Console.WriteLine("Args(): " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString());
                foreach (string xx in args)
                {
                    Console.WriteLine(xx.ToString());

                    switch (xx.ToUpper())
                    {
                        case "PRODUCTS":
                            Helper.Sql_Misc_NonQuery("UPDATE communications..Middleware SET status = 100 WHERE status = 200 AND source_table like 'herroom%'");

                            Products.Process_Brand();
                            Products.Process_Collections();
                            Products.Process_Colors();
                            Products.Process_HerColors();
                            Products.Process_KeywordsFeatures();
                            Products.Process_Product_OST();

                            Products.ReProcess_Product_Configurable_Middleware();
                            Products.ReProcess_Product_Simple_Middleware();

                            Products.Process_Product_Configurable_Middleware();
                            Products.Process_Product_Simple_Middleware();

                            Products.Process_Product_Options_Middleware();

                            Products.ReProcess_Product_Links_Middleware();

                            Products.Generate_Product_Links_Middleware();
                            Products.Process_Product_Links_Middleware();

                            Products.Process_Product_Visibility_Middleware();
                            Products.Process_Product_InStock_Middleware();
                            Helper.Middleware_700_Retry();

                            //Products.Process_Get_Color_Options();

                            break;

                        //  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                        case "PRODUCTCONFIGUPDATES":
                            Products.Process_Configurable_Updates(0, "", "herroom..stylematching", "", 1000);
                            Products.Process_Configurable_Updates(0, "", "herroom..stylefeatures");
                            Products.Process_Configurable_Updates(0, "", "herroom..stylekeywords");
                            Products.Process_Configurable_Updates(0, "", "herroom..stylecategories");
                            Products.Process_Configurable_Updates(0, "", "herroom..productlinks");
                            Products.Process_Configurable_Updates(0, "", "herroom..stylessizingchart");
                            Products.Process_Configurable_Updates(0, "", "herroom..styleproductlinks");
                            Products.Process_Configurable_Updates(0, "", "herroom..stylevideo");
                            Products.Process_Configurable_Updates(0, "", "herroom..stylevideos");
                            Products.Process_Configurable_Updates(0, "", "herroom..stylelabels");
                            Products.Process_Configurable_Updates(0, "", "herroom..stylenewarrival");
                          
                            break;

                        case "PRODUCTLINKS":
                            Products.Process_Configurable_Updates(0, "", "herroom..styleproductlinks", "", 1000);
                            break;

                        case "STYLELABELS":
                            Products.Process_Configurable_Updates(0, "", "herroom..stylelabels", "", 100);
                            //E:\Executables\MagentoAPI-Products-2\bin\MagentoProductAPI.exe STYLELABELS
                            break;

                        case "PRODUCTS-SIMPLE":
                            Products.ReProcess_Product_Simple_Middleware();
                            Products.Process_Product_Simple_Middleware();
                            break;

                        //When a sku color code is changed, it would be best to delete/readd it
                        // but MUST relink to style...
                        case "PRODUCSDELETEREADD":
                            Products.Product_Delete_Readd_Process();
                            break;

                        case "PRODUCTSPRELOAD":
                            //FOR ONCE A DAY RUN: 
                            Helper.Sql_Misc_NonQuery("UPDATE communications..Middleware SET status = 100 WHERE status = 200 AND source_table like 'herroom%'");

                            Products.Process_Get_Color_Options();
                            Products.Process_Brand();
                            Products.Process_Collections();
                            Products.Process_Colors();
                            Products.Process_HerColors();
                            Products.Process_KeywordsFeatures();
                            Products.Process_Product_OST();

                            break;

                        case "PRODUCTSLOADPOST":
                            //Loads POST (NEW) products ONLY
                            Products.Process_Product_Configurable_Middleware(-3);
                            Products.Process_Product_Simple_Middleware(-3);

                            Products.Generate_Product_Links_Middleware();
                            Products.Process_Product_Links_Middleware();

                            //Helper.Sql_Misc_NonQuery("EXEC Communications..proc_mag_product_links_json_api_child 'eilw01-e60032', 0, 0 ");
                            //Products.Process_Product_Links_Middleware("eilw01-e60032", 4);

                            break;

                        case "PRODUCTSLOAD":
                            Products.ReProcess_Product_Configurable_Middleware();
                            //Products.ReProcess_Product_Simple_Middleware();

                            Products.Process_Product_Configurable_Middleware();
                            //Products.Process_Product_Simple_Middleware();  // Seperate task for now 2024-05-16

                            Products.Process_Product_Options_Middleware();

                            Products.ReProcess_Product_Links_Middleware();

                            Products.Generate_Product_Links_Middleware();
                            Products.Process_Product_Links_Middleware();

                            Products.Process_Product_Visibility_Middleware();
                            Products.Process_Product_InStock_Middleware();

                            Products.Process_Item_Price();

                            Helper.Middleware_700_Retry();
                            break;

                        case "PROCESSITEMPRICE":
                            Products.Process_Item_Price(0, "", 500);
                            Products.Process_Style_Price_Bulk(0, 40);
                            break;

                        case "PROCESSITEMPRICESALE":
                            DataTable dtz;
                            dtz = Helper.Sql_Misc_Fetch("SELECT TS.mid, TS.itemid FROM communications..tempsku TS "
                                + " INNER JOIN communications..Middleware MM on MM.id = TS.mid WHERE TS.done = 0 AND MM.status = 100  ORDER BY TS.id ");
                            foreach (DataRow drz in dtz.Rows)
                            {
                                if (Products.Process_Item_Price(long.Parse(drz["mid"].ToString()), "", 5))
                                {
                                    Helper.Sql_Misc_NonQuery("UPDATE communications..tempsku SET done = 1 WHERE itemid = " + drz["itemid"].ToString());
                                }
                                else
                                {
                                    Helper.Sql_Misc_NonQuery("UPDATE communications..tempsku SET done = -1 WHERE itemid = " + drz["itemid"].ToString());
                                }
                            }
                            break;

                        case "PROCESSITEMPRICESALE1":
                            DataTable dtz1;
                            dtz1 = Helper.Sql_Misc_Fetch("SELECT TS.mid, TS.itemid FROM communications..tempsku TS "
                                + " INNER JOIN communications..Middleware MM on MM.id = TS.mid WHERE TS.done = 10 AND MM.status = 100  ORDER BY TS.id ");
                            foreach (DataRow drz in dtz1.Rows)
                            {
                                if (Products.Process_Item_Price(long.Parse(drz["mid"].ToString()), "", 5))
                                {
                                    Helper.Sql_Misc_NonQuery("UPDATE communications..tempsku SET done = 1 WHERE itemid = " + drz["itemid"].ToString());
                                }
                                else
                                {
                                    Helper.Sql_Misc_NonQuery("UPDATE communications..tempsku SET done = -1 WHERE itemid = " + drz["itemid"].ToString());
                                }
                            }
                            break;

                        case "PROCESSITEMPRICESALE2":
                            DataTable dtz2;
                            dtz2 = Helper.Sql_Misc_Fetch("SELECT TS.mid, TS.itemid FROM communications..tempsku TS "
                                + " INNER JOIN communications..Middleware MM on MM.id = TS.mid WHERE TS.done = 20 AND MM.status = 100  ORDER BY TS.id ");
                            foreach (DataRow drz in dtz2.Rows)
                            {
                                if (Products.Process_Item_Price(long.Parse(drz["mid"].ToString()), "", 5))
                                {
                                    Helper.Sql_Misc_NonQuery("UPDATE communications..tempsku SET done = 1 WHERE itemid = " + drz["itemid"].ToString());
                                }
                                else
                                {
                                    Helper.Sql_Misc_NonQuery("UPDATE communications..tempsku SET done = -1 WHERE itemid = " + drz["itemid"].ToString());
                                }
                            }
                            break;

                        case "ITEMSEXPECTEDDATE":
                            Products.Process_Configurable_Updates(0, "", "herroom..itemsexpecteddate", "100", 50);
                            break;

                        case "PRODUCTSMISSING":
                            Products.Process_Products_Active_NotinMagento();
                            break;

                        case "ORDERS":  //  DEPLOYED TO AUTOMATE
                            Orders.Process_Order_Ship();
                            Orders.Process_Order_Invoice();
                            Orders.Process_Order_Comments();
                            Orders.Process_Order_Refund();
                            Orders.Process_Order_Cancel();
                            break;

                        case "ORDERS-MISSINGREPORT":
                            Orders.Missing_Order_Report_API(0, true);
                            break;

                        case "ORDERS_SHIP":
                            Orders.Process_Order_Ship();
                            break;

                        case "ORDERS_INVOICE":
                            Orders.Process_Order_Invoice();
                            //Orders.Missing_Order_Report_API();
                            break;

                        case "ORDERS_COMMENTS":
                            Orders.Process_Order_Comments();
                            break;

                        case "ORDERS_REFUND":
                            Orders.Process_Order_Refund();
                            break;

                        case "ORDERS_CANCEL":
                            Orders.Process_Order_Cancel();
                            break;

                        //DO NOT USE THIS YET
                        case "ORDERS_MISSING_REPORT":
                            long MiddlewareID = 0;

                            MiddlewareID = Orders.Process_Order_Missing_Report();
                            if (MiddlewareID > 0)
                            {
                                //This does the GET on the row 
                                Helper.MagentoApiPush(MiddlewareID);
                                Helper.Sql_Misc_NonQuery("EXEC communications..[proc_mag_report_orders_notinmiddlware]");
                            }
                            break;


                        case "INVENTORY":   // DEPLOYED TO AUTOMATE
                            //Entries created by proc_mag_process_inventory_updates_api
                            Helper.Sql_Misc_NonQuery("EXEC communications..[proc_mag_process_inventory_updates_api]");
                            Products.StockItems_Adjustment();

                            //RUN THIS SO THE SQL JOB WILL PICKUP GET CHANGES;
                            Products.StockItems_Adjustment_Get();
                            break;

                        case "REVIEWAPPROVAL":   // DEPLOYED TO AUTOMATE, run 2x/hr for now
                            Products.Process_Review_Updates_Middleware();
                            Products.Product_Review_Status_Update();

                            //Middleware_Incoming replaced with API GET - 2024-10-09
                            //Products.Process_Reviews_MiddlewareIncoming(0, 2);
                            Products.Product_Reviews_GET();
                            break;

                        case "REVIEWBACKLOG":   // To Catch any stragglers
                            Products.Products_Reviews_2(-30, 10);
                            break;

                        case "CSRKUSTOMERAPI":
                            //CAN ONLY PROCESS 2 at at time without API locking up, investigating
                            Kustomer.Process_Kustomer_Contactus(0, 2, true);
                            //Kustomer.Process_Kustomer_Contactus(0, 2, false);
                            break;

                        case "CSRKUSTOMERAPI-NOEMAIL":
                            //CAN ONLY PROCESS 2 at at time without API locking up, investigating
                            //Kustomer.Process_Kustomer_Contactus(0, 2, true);
                            Kustomer.Process_Kustomer_Contactus(0, 2, false);
                            break;


                        case "CSRKUSTOMEREMAILONLY":
                            Kustomer.Process_Kustomer_Contactus_CSREmailOnly();

                            //2024-08-20 to Pick up HerRoom Question about PDPs, these were coming direct from Magento but are not now
                            Kustomer.Process_Kustomer_Contactus(0, 2, true, false);
                            break;

                        case "KUSTOMERCHECK":
                            DataTable dtK;
                            dtK = Helper.Sql_Misc_Fetch("SELECT TOP 25 ID FROM communications..magento_contactus with (nolock) WHERE len(ISNULL(conversationid, '')) = 0 AND Comment255 IS NOT NULL AND emailsent IN (2,3,200,600) AND Datestamp between Dateadd(minute, -90, Getdate()) AND Dateadd(minute, -5, Getdate())  order by 1 ");
                            foreach (DataRow drK in dtK.Rows)
                            {
                                Kustomer.TestforContactusinKustomer(long.Parse(drK["id"].ToString()));
                            }
                            break;

                        //2024-10-16
                        case "KUSTOMERAPI":
                            Customers.PDPQuestion_Fetch_Process(0, -1);
                            Customers.Contactus_Fetch_Process(0, "", -1);

                            Kustomer.Process_Kustomer_Contactus(0, 2, true, false);

                            break;

                        case "ORDERINVOICEPAYPAL":
                            Orders.PayPal_Invoice_Order_Update();
                            break;

                        case "IVR":
                            //GETS Magento and HT INVENTORY AND COMPARES, PUSHES RESULTS TO Communications..Inventory_Issues Table
                            Inventory.Process_Low_Inventory();

                            //THIS EMAILS IVR 
                            Helper.Sql_Misc_NonQuery("EXEC Communications..[proc_mag_report_inventory_issues]");

                            //THIS CORRECTS STOCK ISSUES
                            Helper.Sql_Misc_NonQuery("EXEC communications..[proc_mag_inventory_issues_process]");

                            //Runs Inventory Sync to process new Middleware Rows created to resolve issues;
                            Products.StockItems_Adjustment();

                            break;

                        case "STAGING": // NO LONGER IN USE 
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET LoadStaging=700 WHERE LoadStaging = 100 AND status = 700");
                            Staging.Process_Middleware_Staging(0, "herroom..manufacturers");
                            Staging.Process_Middleware_Staging(0, "herroom..collections");
                            Staging.Process_Middleware_Staging(0, "herroom..colors");
                            Staging.Process_Middleware_Staging(0, "her_color");
                            Staging.Process_Middleware_Staging(0, "herroom..keywords");
                            Staging.Process_Middleware_Staging(0, "OtherSearchTerms");
                            Staging.Process_Middleware_Staging(0, "herroom..styles");
                            Staging.Process_Middleware_Staging(0, "herroom..items");
                            Staging.Process_Middleware_Staging(0, "herroom..styleOptions");
                            Staging.Process_Middleware_Staging(0, "Herroom..ItemsLink");
                            Staging.Process_Middleware_Staging(0, "herroom..styleVisibility");
                            break;


                        case "RELOAD":  // DB product reload - hopefully will never need again
                            //Products.Process_Product_Configurable_Middleware();
                            //products.Process_Product_Simple_Middleware(); 
                            Products.Process_Product_Configurable_Middleware(-1);
                            //Products.Process_Product_Options_Middleware();
                            //Products.Process_Product_Simple_Middleware(0, 101);
                            //Products.Process_Product_Links_Middleware();
                            //Products.Generate_Product_Links_Middleware();
                            //Products.Process_Product_Links_Middleware();
                            // Products.Process_Product_Visibility_Middleware();
                            break;

                        case "DBTEST":   
                            //"MagentoProductAPI.exe DBTEST"  
                            //E:\Executables\MagentoAPI\bin\MagentoProductAPI.exe DBTEST    
                            //Orders.Order_Comments_DB();
                            Helper.Sql_Misc_NonQuery("INSERT herroom..SQLLog(Log_DTM, SP_Name, Log_Text) SELECT Getdate(), 'TEST', 'FROM API'");
                            break;

                        case "APIPUSH":
                            if (Arg2.Length > 0 && long.TryParse(Arg2, out Arg2Int))
                            {
                                Helper.MagentoApiPush(Arg2Int);
                            }
                            break;

                        case "PRODUCTCHILDRENFIX":
                            Products.Product_Children(0, 100);
                            break;

                        case "PROCESSPRODUCTLINKS":     //default is 300 rows
                            Products.Process_Product_Links_Middleware("", 200);
                            break;

                        case "PROCESSPRODUCTLINKS10":  //default is 300 rows
                            Products.Process_Product_Links_Middleware("", 10);
                            break;

                            // NO LONGER IN USE 2024-10-01
                        case "NEWRMAPROCESS":
                            DataTable dtRMA;
                            dtRMA = Helper.Sql_Misc_Fetch("EXEC Communications..proc_mag_insert_middleware_get_newrma_row;");
                            if (dtRMA.Rows.Count == 1)
                            {
                                Helper.MagentoApiPush(long.Parse(dtRMA.Rows[0]["maxid"].ToString()));

                                RMA.Process_Incoming_RMA(long.Parse(dtRMA.Rows[0]["maxid"].ToString()));
                            }
                            break;

                        // NO LONGER IN USE 2024-10-01
                        case "RMASHIPLABELAPI":
                            RMA.ReturnLabelPDF_API_Process();
                            break;

                        // NEW 2024-10-01
                        case "RMAFETCHPROCESS":
                            RMA.Fetch_RMA_MiddlewareIncomingInsert();
                            RMA.Process_RMA_MiddlewareIncomingTable(50);
                            break;

                        case "EMAILTEST":  //TO TEST EMAIL GETTING SENT FROM VARIOUS SERVERS
                            Helper.SendEmail("HisRoom - Order Status", "Website: HisRoom <br />Sender Name: Brogan <br />Email: broganray@yahoo.com <br />Type: Order Status <br />Order no: 0<br />Brand: <br />Band Size: <br />Cup Size: <br />Customer BraSizes:  <br />Country: US<br />Comment: Where is my order?! <br />C: 6427737 <br />O: 10774143D-  <br />PQ: 1 <br />BrowserInfo: Country: <br />user_agent: Mozilla/5.0 (iPhone; CPU iPhone OS 17_3_1 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.3.1 Mobile/15E148 Safari/604.1<br />browser_name: Safari<br />browser_version: 605.1.15<br />ip_address: 104.28.50.198<br /> <br />", "csr-site@herroom.com", "Brogan<broganray@yahoo.com>", "", "Customer Service<customerservice@hisroom.com>;Amber<amber@herroom.com>;Windy<windy@herroom.com>;Thomas Tribble<Thomas@andragroup.com>");
                            break;

                        case "INVOICEMANUALPROCESS":
                            long RowtoGet = 0;
                            DataTable dtInv = Helper.Sql_Misc_Fetch("EXEC communications..[proc_mag_orderinvoice_paypal_htorder_getinsert] @daysback = -1");
                            if (dtInv.Rows.Count == 1)
                            {
                                RowtoGet = long.Parse(dtInv.Rows[0][0].ToString());
                                Console.WriteLine("RowtoGet: " + RowtoGet.ToString());
                                Helper.MagentoApiPush(RowtoGet);
                                Orders.Invoice_Order_Get_Update(RowtoGet);
                            }
                            break;

                        case "FULFILLMENTDATEBULK":
                            Products.Process_Product_Simple_Fulfillmentdate_Bulk_DailyRun();
                            break;

                        case "FULFILLMENTDATEUPDATES":
                            //MagentoProductAPI.exe FULFILLMENTDATEUPDATES  
                            Products.Process_Product_Simple_Fulfillmentdate_Bulk();
                            break;


                        case "ASPTEST":
                            Helper.Sql_Misc_NonQuery("INSERT Communications..sqllog(process, event, Datestamp, Processed) SELECT 'ASP TEST', 'TEST', getdate(), 0");
                            break;

                        case "URLLOAD":
                            // "e:\Executables\MagentoAPI-Products-2\bin\MagentoProductAPI.exe URLLOAD"  
                            long Counter = 0;
                            DataTable dtUrl;
                            dtUrl = Helper.Sql_Misc_Fetch("SELECT ID FROM Communications..Middleware with (nolock) WHERE worker_id = 2000 AND Status = 100 order by id ");
                            foreach (DataRow drUrl in dtUrl.Rows)
                            {
                                //Products.Process_Product_Simple_ItemsUrlKey(long.Parse(drUrl["id"].ToString()), 0, "mstaging");
                                Products.Process_Product_Simple_ItemsUrlKey(long.Parse(drUrl["id"].ToString()));
                                Counter++;

                                if (Counter % 100.0 == 0)
                                {
                                    Console.WriteLine(Counter.ToString());
                                }
                            }
                            break;

                        case "URLLOAD3000":
                            // "e:\Executables\MagentoAPI-Products-2\bin\MagentoProductAPI.exe URLLOAD"  
                            long Counter3 = 0;
                            DataTable dtUrl3;
                            dtUrl3 = Helper.Sql_Misc_Fetch("SELECT ID FROM Communications..Middleware with (nolock) WHERE worker_id = 3000 AND Status = 100 order by id ");
                            foreach (DataRow drUrl3 in dtUrl3.Rows)
                            {
                                //Products.Process_Product_Simple_ItemsUrlKey(long.Parse(drUrl["id"].ToString()), 0, "mstaging");
                                Products.Process_Product_Simple_ItemsUrlKey(long.Parse(drUrl3["id"].ToString()));
                                Counter3++;

                                if (Counter3 % 100.0 == 0)
                                {
                                    Console.WriteLine(Counter3.ToString());
                                }
                            }
                            break;


                        case "URLLOADSTAGING":
                            // "e:\Executables\MagentoAPI-Products-2\bin\MagentoProductAPI.exe URLLOAD"  
                            long CounterS = 0;
                            DataTable dtUrls;
                            dtUrls = Helper.Sql_Misc_Fetch("SELECT  ID FROM Communications..Middleware with (nolock) WHERE worker_id = 2000 AND Status = 100 AND ID >= 14749574 order by id ");
                            foreach (DataRow drUrl in dtUrls.Rows)
                            {
                                Products.Process_Product_Simple_ItemsUrlKey(long.Parse(drUrl["id"].ToString()), 0, "mstaging");
                                CounterS++;

                                if (CounterS % 100.0 == 0)
                                {
                                    Console.WriteLine(CounterS.ToString());
                                }
                            }
                            break;

                        case "KLAVIYOSISTERHOOD":
                            Klaviyo.SisterhoodBras_API();
                            Klaviyo.SisterhoodPanty_API();
                            break;

                        case "KLAVIYOFULFILLED":
                            Klaviyo.Fulfilled_API("Fulfilled");
                            break;

                        case "KLAVIYOORDERSYNC":
                            Klaviyo.Fulfilled_API("Placed", 60);
                            Klaviyo.Fulfilled_API("Cancelled", 60);
                            Klaviyo.Fulfilled_API("Returned", 60);
                            break;

                        case "RMAPROCESSING":
                            RMA.Process_Incoming_DirectRMA_Insert_2();
                            break;

                        case "PROPROCESS":
                            if (Arg2.Length > 0 && long.TryParse(Arg2, out Arg2Int))
                            {
                                //Console.WriteLine(Arg2Int.ToString());
                               //DevMode = -1;
                                POHerTools.PO_HT_Process(long.Parse(Arg2Int.ToString()));
                                //Helper.Sql_Misc_NonQuery("INSERT herroom..SQLLog(Log_DTM, SP_Name, Log_Text) SELECT Getdate(), 'TEST', 'FROM API : " + Arg2 + "'");
                            }
                            else
                            {
                                POHerTools.PO_HT_Process();
                            }
                            break;

                        case "POPROCESSPEACHTREE":
                            POHerTools.PO_PeachTree_Report();   
                            break;

                        case "PO850FTPRECORD":
                            POProcess.Incoming_850_FTP_Record();
                            break;

                        case "856FILEPROCESS":
                            // "MagentoProductAPI.exe 856FILEPROCESS"  
                            //POHerTools.Incoming_856File_Process();
                            POProcess.Incoming_856File_Process();
                            POProcess.ASN_Daily_Report_Process();
                            break;

                        case "856FILEPROCESSNOFILEMOVE":
                            POProcess.Incoming_856File_Process(true);
                            break;

                        case "856FILEPROCESSMOVE":
                            POProcess.Incoming_856File_FTP_Move();
                            break;
                        // E:\Executables\MagentoAPI\bin\MagentoProductAPI.exe 856FILEPROCESSNOFILEMOVE

                        //one off processing
                        case "856FILEPROCESSASNREPORT":
                            POProcess.ASN_Daily_Report_Process();
                            // E:\Executables\MagentoAPI\bin\MagentoProductAPI.exe 856FILEPROCESSASNREPORT 
                            break;

                            //one off processing
                        case "856FILEPROCESSXXX":
                            string PONumber;
                            //POProcess.Incoming_856File(@"\\task\Automate\Working\PO-SPS\Inbound\SH_52709717066.txt", "SH_52709717066.txt", out PONumber);
                            
                            // E:\Executables\MagentoAPI\bin\MagentoProductAPI.exe 856FILEPROCESSXXX  
                            break;

                        case "ABANDONEDCARTSREPORT":
                            // proc_mag_abandondedcarts_report_weekly is run from a SQL job
                            Orders.AbandondedCartsReport_Process();
                            Helper.Sql_Misc_NonQuery("EXEC communications..[proc_mag_abandondedcarts_report]");
                            break;

                        case "ORDERINVOICEFETCH":
                            Orders.OrderInvoiceFetchProcessing();
                            break;

                        case "LABELTEST":
                            // E:\Executables\MagentoAPI\bin\MagentoProductAPI.exe LABELTEST   
                            DevMode = 2;
                            //POHerTools.PO_PDFLabels_Send("Shdw01", "x1000");
                            //POHerTools.TESTEMAIL("HAY", @"\\tools\e$\www\barcodeGen3\PDF_Files\Shdw01\test.pdf");
                            break;

                        case "SPSREPORTREADER" :
                            POProcess.SPS_Report_Reader("");
                            //RUN REPORT
                            Helper.Sql_Misc_NonQuery("EXEC Communications..[proc_mag_report_sps_ftplog]");
                            break;

                        case "SPSREPORTREADERWEEKLY":
                            //RUN REPORT
                            Helper.Sql_Misc_NonQuery("proc_mag_report_sps_ftplog_weekly]");
                            break;


                        //sproc proc_mag_refunds_processing runs at 6:10am
                        case "MAGENTOREFUNDSAPI":
                            DataTable dtO;
                            dtO = Helper.Sql_Misc_Fetch("SELECT top 500 * FROM (SELECT distinct orderno [orderno] from Magento_Refunds_API where DateUpdated IS NULL) xx  order by 1");
                            foreach (DataRow drO in dtO.Rows)
                            {
                                Orders.OrderRefundAPI(int.Parse(drO["orderno"].ToString()), 0);
                            }
                            break;
                    }
                }
            }
            else
            {
                DevMode = 1;
                // ALL OF THIS IS TO RUN FUNCTIONS MANUALLY AS NEEDED 

                //Customers.PDPQuestion_Fetch_Process(39787, -1);

                //Customers.PDPQuestion_Fetch_Process(0, -1);
                //Customers.Contactus_Fetch_Process(0, "", -1);

                //Kustomer.Process_Kustomer_Contactus(27278, 2, false, true);

             
                //2024-10-29 WORKS - DAILY JOB !! 

                //Orders.OrderRefundAPI(10607097, 0);
                /*
                DataTable dtO;
                dtO = Helper.Sql_Misc_Fetch("SELECT top 500 * FROM (SELECT distinct orderno [orderno] from Magento_Refunds_API where DateUpdated IS NULL) xx  order by 1");
                foreach(DataRow drO in dtO.Rows)
                {
                    Orders.OrderRefundAPI(int.Parse(drO["orderno"].ToString()), 0);
                }
                */


                //string Kustomerid;
                //string Convoid;
                //DataTable DTNote;
                //Kustomerid = Kustomer.GetKustomerid("cluksicfamily@gmail.com");
                //if (Kustomerid.Length > 5 && Kustomerid.Substring(0,5) != "ERROR")
                //{

                //} 

                //Customers.Contactus_Fetch_Process(158, "", 0);

                //POProcess.Incoming_850_FTP_Record();

                //2024-10-24
                //POProcess.SPS_Report_Reader("");

                //2024-10-21
                //Kustomer.PushKustomerOrderNotes(26594);




                //Kustomer.PushKustomerOrderNotes(26403);


                //Customers.PDPQuestion_Fetch_Process(0, -1);
                //Customers.Contactus_Fetch_Process(0, "", -1);
                //Kustomer.Process_Kustomer_Contactus(26439, 12, true, false);

                //Kustomer.Process_Kustomer_Contactus(26664, 2, false, true);

                //Kustomer.Process_Kustomer_Contactus(26634, 2, false, true);

                //Products.Process_Review_Updates_Middleware();

                //Products.Product_Reviews_GET();

                //Helper.MagentoApiPush(16578622);

                //Products.Product_Children(0, 10);

                //2024-10-16
                //Customers.Contactus_Fetch_Process(8, "", 0);
                //Customers.Contactus_Fetch_Process(0, "", -1);  // KEEP TESTING

                //Kustomer.TestforContactusinKustomer(26171);

                //Customers.PDPQuestion_Fetch_Process(39773, 0);

                //Kustomer.Process_Kustomer_Contactus(0, 2, true, false);

                //Kustomer.Process_Kustomer_Contactus(26345, 1, false, true);

                //Customers.PDPQuestion_Fetch_Process(0, -1);


                //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                //2024-10-16
                /*
                DataTable dtx;
                dtx = Helper.Sql_Misc_Fetch("select top 5000 Magentoorderno FROM Hercust..orders where magentoorderno > '2000296104' AND cancel = 0 and shipped = 0 and siteversion = 1 AND Timestamp > '2024-06-01' order by 1 ");
                foreach (DataRow drx in dtx.Rows)
                {
                    Orders.OrderSumCheck(drx["magentoOrderno"].ToString());
                }
                */

                //Orders.Process_Order_Invoice(16569033);

                // !!!
                //RMA.Process_RMA_MiddlewareIncomingTable(50);

                //POProcess.PO_HT_Process(16512804);

                //POProcess.Incoming_856File_FTP_Move();

                // Orders.OrderInvoiceFetch("001052519", "528173", 16491674);

                //Orders.OrderInvoiceFetch("2000297066", "528418");

                // !!!!!!!

                //Orders.OrderInvoiceFetchProcessing();               

                //Orders.OrderSumCheck("2000296600");


                Console.WriteLine(" ------------");



                //string POnumber;
                //POProcess.Incoming_856File("C:\\Automate\\Working\\PO-SPS\\Inbound\\SH_52464123788.txt", "SH_52464123788", out POnumber );

                //POHerTools.Incoming_856File_Process();

                // !!! SHOULD WORK 
                //POHerTools.Incoming_856File_Process();
                //POProcess.ASN_Daily_Report_Process();



                //POProcess.ASN_Daily_Report_Process("2024-09-13");

                //POProcess.ASN_Daily_Report_Process();

                //Helper.MagentoApiPush(16289001);  /////////////////////////////////////////////

                // Products.Process_Configurable_Updates(16195548, "", "herroom..stylematching", "", 100);

                //Orders.Process_Order_Invoice();
                //Orders.Process_Order_Ship(16181158);
                //Orders.Process_Order_Comments(16181159);


                //Orders.Process_Order_Refund(16240508);

                // Products.Process_HerColors(16187451);

                //Products.Product_Review_Status_Update(357808);

                //POHerTools.PO_PDFLabels_Send("pj001", "216818");

                //Products.Process_Configurable_Updates(16172975, "", "herroom..stylematching", "", 2);

                //Products.Process_Configurable_Updates(0, "", "herroom..stylematching", "", 1000);
                //POProcess.POCreate("cal001", "x2000");

                //Products.Process_Product_Configurable_Middleware(16683353);

                //POHerTools.TESTEMAIL("HAY", @"\\tools\e$\www\barcodeGen3\PDF_Files\Shdw01\test.pdf");

                //Orders.AbandondedCartsReport_Process();
                //Orders.AbandondedCartsReport_Fetch(15934088);

                //POHerTools.PO_PDFLabels_Send("Shdw01", "x1000");

                //Orders.PayPal_Invoice_Order_Update();

                //RMA.Process_Incoming_DirectRMA_Insert_2(0);

                //Products.Process_Item_Price(0, "", 3);

                //Products.Process_Style_Price_Bulk(15675388);



                //POHerTools.Incoming_856File_Process();
                //POHerTools.Incoming_856File(@"C:\Automate\Working\PO-SPS\Inbound\SH_51178073500.txt", "SH_51178073500");


                //POHerTools.Incoming_856File(@"C:\Users\thomas\Documents\Reference\SPS\Incoming\SH_51462778051.txt", "SH_51462778051x");

                //POProcess.PODropShipCreate("8754182", "wac001");

                //POHerTools.PO_HT_Process(1588421800);

                //POHerTools.PO_PDFLabels_Send("pj001"); 

                //Orders.Process_Order_Ship(15826243);


                //POHerTools.PO_PeachTree_Report();  // WORKS 

                //Products.Process_Configurable_Updates(15659461, "", "herroom..stylevideos");


                //Helper.MagentoApiPush(16656028);
                //Orders.PayPal_Invoice_Order_Update(16138497);



                //Helper.MagentoApiPush(16293822);  /////////////////////////////////////////////
                //Orders.Invoice_Order_Get_Update(15815836); 
                //Helper.Sql_Misc_NonQuery("EXEC communications..[proc_mag_orderinvoice_all_htorder_update] @GETid = 14939719");

                //Products.Process_Style_Price_Bulk(14879259);   //eileen west 2024-06-04

                //Products.Process_Product_Simple_ItemsUrlKey(14568101);

                // Products.Process_Item_Price(0, "", 300);

                //Products.Process_Product_Simple_ItemsUrlKey(14790495);


                // Products.StockItems_Adjustment(16254969);
                // Products.StockItems_Adjustment(16254968);

                //Products.Process_Configurable_Updates(0, "", "herroom..itemsexpecteddate", "100", 50);
                // Products.Process_Product_Configurable_Middleware(16684023);


                //Products.Process_Product_Configurable_Middleware(16708523);


                //Products.Process_Product_Simple_Middleware(0, 100, 100);

                //Products.Process_Product_Simple_Middleware(16237554);
                //Products.Process_Product_Simple_Middleware(16237555);


                ///////////////////////////////////////////////////////////////////////////////////
                //Products.Process_Item_Price(0, "Fre001-AS0056", 500);
                // Products.Process_Item_Price(0, "", 350);
                //Products.Process_Style_Price_Bulk(xx);

                //Helper.MagentoApiPush(14842313);

                //Helper.MagentoApiPush(14898682);
                //Orders.Invoice_Order_Get_Update(14488844);

                // Orders.Process_Order_Invoice();



                //Products.Process_Product_Simple_Middleware(0);

                // !!!!!!!!!!!!!!! USE THIS - Style/Sku RELOAD 

                //Products.Process_Product_Configurable_Middleware(16417733);

                //Products.Process_Product_Simple_Middleware(xx);

                // !!!!

                //Products.Process_Product_Configurable_Middleware(16738085);

                //Products.Process_Product_Options_Middleware(16717612);
                //Products.Process_Product_Options_Middleware(16717611);


                //Products.Process_Product_Simple_Middleware(16329012);

                //Helper.Sql_Misc_NonQuery("EXEC Communications..proc_mag_product_links_json_api_child 'Prd01-056-3490', 1, 0 ");
                // Products.Process_Product_Links_Middleware("Prd01-056-3490", 2);


                //   Products.Process_Reviews_MiddlewareIncoming(3378765, 2);
                // Products.Process_Reviews_MiddlewareIncoming(3378545, 2);

                //Products.Process_Style_Price_Bulk(13914936);

                //Products.Process_Item_Price(0, "", 249);

                //Products.Process_Item_Price(13778765);

                //Products.Process_Item_Price(0, "", 25);
                //Products.Process_Style_Price_Bulk(0, 5);


                //Products.Process_Product_Simple_Fulfillmentdate_Bulk(16544581);

                //Helper.MagentoApiPush(16656028, "mstaging");

                Console.WriteLine("END"); 
                Console.WriteLine("----");

            }      
        }
       
    }

}
