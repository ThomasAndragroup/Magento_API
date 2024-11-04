using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data;
using Newtonsoft.Json;
using System.Diagnostics;
using System.Collections;

namespace MagentoProductAPI
{
    class Products
    {

        //keywords enpoint_name; https://mcprod.herroom.com/bulk/V1/products/attributes/ag_keywords/option

        //Style/SKU updates are posted directly into Middleware and PreProcessor Tables.
        //Do Middleware processing first for all related stuff THEN Styles/SKU
        //Then do preprocessor Rows, color to skus step by step

        public static Boolean Process_Colors(long Middlewareid = 0)
        {
            Boolean bReturnvalue = true;
            DataTable dt = new DataTable();
            String APIReturn = "";
            long mid;

            try
            {
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND source_table = 'herroom..colors' AND id = " + Middlewareid.ToString());
                }
                else
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND source_table = 'herroom..colors' ORDER BY ID");
                }

                Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                {
                    APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");

                    if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                    {
                        //Middleware should already by set to 500
                        Console.WriteLine("ERROR for Color Processing: " + dr["source_id"].ToString());
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, from_magento = '" + APIReturn + "' WHERE ID = " + dr["id"].ToString());

                        Helper.SendEmail("ERROR for Color Processing: " + dr["source_id"].ToString(), "ERROR for Color Processing: " + dr["source_id"].ToString() + "; ID: " + dr["id"].ToString(), "thomas@andragroup.com", "alerts@herroom.com");
                    }
                    else //Should return new value
                    {
                        if (long.TryParse(APIReturn, out mid) && dr["endpoint_method"].ToString().ToUpper() == "POST")
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + mid.ToString() + "' WHERE ID = " + dr["id"].ToString());

                            //do THIS FOR ALL and update sprocs to use BatchID and set batchnum = null
                            if (mid > 0)
                            {
                                Helper.Sql_Misc_NonQuery("Update Herroom..Colors SET Magentoid = " + mid.ToString() + " WHERE colorcode = '" + dr["source_id"].ToString() + "' AND Magentoid IS NULL");
                                //UPDATE ONLY ITEMS WITH THAT COLORCODE 
                                Helper.Sql_Misc_NonQuery("UPDATE communications..Middleware SET Batchnum = 100 WHERE status=100 AND ISNULL(Batchid,'x') = '" + dr["batchid"].ToString() + "' AND source_table = 'herroom..items'; ");
                            }
                        }
                        else
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + APIReturn + "' WHERE ID = " + dr["id"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Process_Color(" + Middlewareid.ToString() + ": " + ex.ToString());
                bReturnvalue = false;
            }

            return bReturnvalue;
        }

        public static Boolean Process_HerSizes(long Middlewareid = 0)
        {
            Boolean bReturnvalue = true;
            DataTable dt = new DataTable();
            String APIReturn = "";
            long mid;

            try
            {
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE source_table = 'her_size' AND id = " + Middlewareid.ToString());
                }
                else
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND source_table = 'her_size' ORDER BY ID");
               }

                Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                {
                    APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");

                    if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                    {
                        //Middleware should already by set to 500
                        Console.WriteLine("ERROR for Her_Size Processing: " + dr["source_id"].ToString());
                    }
                    else //Should return new value
                    {
                        if (long.TryParse(APIReturn, out mid) && dr["endpoint_method"].ToString().ToUpper() == "POST")
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + mid.ToString() + "' WHERE ID = " + dr["id"].ToString());
                            Helper.Sql_Misc_NonQuery("Update Herroom..MasterSize SET Hersizeid = " + mid.ToString() + " WHERE size = '" + dr["source_id"].ToString() + "' AND Hersizeid IS NULL");
                            //UPDATE ONLY ITEMS WITH THAT Her_Size 
                           //Helper.Sql_Misc_NonQuery("UPDATE communications..Middleware SET Batchnum = 100 WHERE status=100 AND ISNULL(Batchid,'x') = '" + dr["batchid"].ToString() + "' AND source_table = 'herroom..items'; ");
                        }
                        else
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + APIReturn + "' WHERE ID = " + dr["id"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Process_HerSizes(" + Middlewareid.ToString() + ": " + ex.ToString());
                bReturnvalue = false;
            }

            return bReturnvalue;
        }


        public static Boolean Process_HerColors(long Middlewareid = 0)
        {
            Boolean bReturnvalue = true;
            DataTable dt ;
            String APIReturn = "";
            long mid;

            try
            {
                if (Middlewareid > 0)
                {
                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("SELECT id, source_id, endpoint_method, ISNULL(batchid, '')[batchid] FROM Communications..Middleware with(nolock) WHERE source_table = 'her_color' AND id = " + Middlewareid.ToString());
                    }
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE source_table = 'her_color' AND id = " + Middlewareid.ToString());
                }
                else
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND source_table = 'her_color' ORDER BY ID");
                }

                Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                {
                    APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");

                    if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                    {
                        //Middleware should already by set to 500
                        Console.WriteLine("ERROR for Color Processing: " + dr["source_id"].ToString());
                    }
                    else //Should return new value
                    {
                        if (long.TryParse(APIReturn, out mid) && dr["endpoint_method"].ToString().ToUpper() == "POST")
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + mid.ToString() + "' WHERE ID = " + dr["id"].ToString());
                            Helper.Sql_Misc_NonQuery("Update Herroom..Colors SET Hercolorid = " + mid.ToString() + " WHERE colorcode = '" + dr["source_id"].ToString() + "' AND Hercolorid IS NULL");
                            //UPDATE ONLY ITEMS WITH THAT Her_COLORCODE 
                            Helper.Sql_Misc_NonQuery("UPDATE communications..Middleware SET Batchnum = 100 WHERE status=100 AND ISNULL(Batchid,'x') = '" + dr["batchid"].ToString() + "' AND source_table = 'herroom..items'; ");

                        }
                        else
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + APIReturn + "' WHERE ID = " + dr["id"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Process_HerColor(" + Middlewareid.ToString() + ": " + ex.ToString());
                bReturnvalue = false;
            }

            return bReturnvalue;
        }

        // Updates style collection only 
        // Entries generated by [proc_mag_product_configurable_json_collections_api]  
        //  Middleware source_table: 'herroom..stylecollection'
        public static Boolean Process_Style_Collections(long Middlewareid = 0, int MaxRowstoProcess = 200)
        {
            Boolean bReturnvalue = true;
            DataTable dt = new DataTable();
            String APIReturn = "";

            try
            {
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE source_table = 'herroom..stylecollection' AND id = " + Middlewareid.ToString());
                }
                else
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT TOP " + MaxRowstoProcess.ToString() + " id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND source_table = 'herroom..stylecollection' ORDER BY Posted");
                }

                Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                {
                    APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Process_Style_Collections(" + Middlewareid.ToString() + ": " + ex.ToString());
                bReturnvalue = false;
            }

            return bReturnvalue;
        }

        /// <summary>
        /// Updated 2024-02-19 for new ag_collection attribute
        /// new sproc for creating entries: [proc_mag_json_collection_name_api]
        /// </summary>
        /// <param name="Middlewareid"></param>
        /// <returns></returns>
        public static Boolean Process_Collections(long Middlewareid = 0)
        {
            Boolean bReturnvalue = true;
            DataTable dt = new DataTable();
            String APIReturn = "";
            long mid;

            try
            {
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND source_table = 'herroom..collections' AND id = " + Middlewareid.ToString());
                }
                else
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND source_table = 'herroom..collections' ORDER BY ID");
                }

                Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                {
                    APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");

                    if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                    {
                        //Middleware should already by set to 500
                        Console.WriteLine("ERROR for Collection Processing: " + dr["source_id"].ToString());
                    }
                    else //Should return new value
                    {
                        if (long.TryParse(APIReturn, out mid) && dr["endpoint_method"].ToString().ToUpper() == "POST")
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + mid.ToString() + "' WHERE ID = " + dr["id"].ToString());
                            // 2024-02-19 New ag_collection attribute
                            Helper.Sql_Misc_NonQuery("UPDATE Herroom..collections SET Magentoid = " + mid.ToString() + " WHERE collectionname = '" + dr["source_id"].ToString() + "' AND ISNULL(Magentoid,0) = 0");
                            
                            //UPDATE Styles for Json Update
                            Helper.Sql_Misc_NonQuery("UPDATE communications..Middleware SET Batchnum = 100 WHERE status=100 AND source_id = '" + dr["batchid"].ToString() + "' AND source_table = 'herroom..styles'; ");
                        }
                        else
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + APIReturn + "' WHERE ID = " + dr["id"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Process_Collections(" + Middlewareid.ToString() + ": " + ex.ToString());
                bReturnvalue = false;
            }

            return bReturnvalue;
        }

        public static Boolean Process_KeywordsFeatures(long Middlewareid = 0)
        {
            Boolean bReturnvalue = true;
            DataTable dt = new DataTable();
            String APIReturn = "";
            long mid;

            try
            {
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(Batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND source_table IN ('herroom..keywords','herroom..features') AND id = " + Middlewareid.ToString());
                }
                else
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(Batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND source_table IN ('herroom..keywords','herroom..features') ORDER BY ID");
                }

                Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                {
                    APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");

                    if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                    {
                        //Middleware should already by set to 500
                        Console.WriteLine("ERROR for KeywordFeature Processing: " + dr["source_id"].ToString());
                    }
                    else //Should return new value
                    {
                        if (long.TryParse(APIReturn, out mid) && dr["endpoint_method"].ToString().ToUpper() == "POST")
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + mid.ToString() + "' WHERE ID = " + dr["id"].ToString());
                            Helper.Sql_Misc_NonQuery("UPDATE Herroom..Keywords SET Magentoid = " + mid.ToString() + " WHERE keywordid = '" + dr["source_id"].ToString() + "' AND Magentoid IS NULL");
                            //UPDATE Styles for Json Update
                            Helper.Sql_Misc_NonQuery("UPDATE communications..Middleware SET Batchnum = 100 WHERE status=100 AND source_id = '" + dr["batchid"].ToString() + "' AND source_table = 'herroom..styles'; ");
                        }
                        else
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + APIReturn.Replace("'", "''") + "' WHERE ID = " + dr["id"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Process_KeywordsFeatures(" + Middlewareid.ToString() + ": " + ex.ToString());
                bReturnvalue = false;
            }

            return bReturnvalue;
        }

        public static Boolean Process_Product_OST(long Middlewareid = 0, int MaxRowstoProcess = 1000)
        {
            Boolean bReturnvalue = true;
            DataTable dt;
            String APIReturn = "";
            long mid;

            try
            {
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,0) [batchid] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND source_table = 'OtherSearchTerms' AND id = " + Middlewareid.ToString());
                }
                else
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT TOP " + MaxRowstoProcess.ToString() + " id, source_id, endpoint_method, ISNULL(batchid,0) [batchid] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND source_table = 'OtherSearchTerms' ORDER BY ID");
                }

                Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                {
                    APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");

                    if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                    {
                        //Middleware should already by set to 500
                        Console.WriteLine("ERROR for OST Processing: " + dr["source_id"].ToString());
                    }
                    else //Should return new value
                    {
                        if (long.TryParse(APIReturn, out mid) && dr["endpoint_method"].ToString().ToUpper() == "POST")
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + mid.ToString() + "' WHERE ID = " + dr["id"].ToString());
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..OtherSearchTerms SET Magentoid = " + mid.ToString() + " WHERE OtherSearchTerm = '" + dr["source_id"].ToString() + "' AND Magentoid IS NULL");
                            //UPDATE Styles for Json Update
                            Helper.Sql_Misc_NonQuery("UPDATE communications..Middleware SET Batchnum = 100 WHERE status=100 AND source_id = '" + dr["batchid"].ToString() + "' AND source_table = 'herroom..styles'; ");
                        }
                        else
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + APIReturn.Replace("'", "''") + "' WHERE ID = " + dr["id"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Process_Product_OST(" + Middlewareid.ToString() + ": " + ex.ToString());
                bReturnvalue = false;
            }

            return bReturnvalue;
        }


        public static Boolean Process_Brand(long Middlewareid = 0)
        {
            Boolean bReturnvalue = true;
            DataTable dt = new DataTable();
            String APIReturn = "";
            long mid;

            try
            {
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE source_table = 'herroom..manufacturers' AND id = " + Middlewareid.ToString());
                }
                else
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND source_table = 'herroom..manufacturers' ORDER BY ID ");
                }

                Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                {
                    APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");

                    if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                    {
                        //Middleware should already by set to 500
                        Console.WriteLine("ERROR for Manufacturers Processing: " + dr["source_id"].ToString());
                    }
                    else //Should return new value
                    {
                        if (long.TryParse(APIReturn, out mid) && dr["endpoint_method"].ToString().ToUpper() == "POST")
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + mid.ToString() + "' WHERE ID = " + dr["id"].ToString());
                            Helper.Sql_Misc_NonQuery("UPDATE Herroom..Manufacturers SET Magentoid = " + mid.ToString() + " WHERE mfrid = '" + dr["source_id"].ToString() + "' AND Magentoid IS NULL");
                            //UPDATE Styles for Json Update
                            Helper.Sql_Misc_NonQuery("UPDATE communications..Middleware SET Batchnum = 100 WHERE status = 100 AND source_id = '" + dr["batchid"].ToString() + "' AND source_table = 'herroom..styles'; ");
                        }
                        else
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + APIReturn.Replace("'", "''") + "' WHERE ID = " + dr["id"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Process_Brand(" + Middlewareid.ToString() + ": " + ex.ToString());
                bReturnvalue = false;
            }

            return bReturnvalue;
        }

        public static Boolean Process_Product_Simple_Fulfillmentdate_Bulk_DailyRun()
        {
            Boolean Returnvalue = true;

            try
            {
                //Create Entries to Process
                Helper.Sql_Misc_NonQuery("EXEC Communications..[proc_mag_products_simple_fulfillmentdate_bulk_api] @LimitProcessing = 0");

                //RUN updates
                //Returnvalue = Process_Product_Simple_Fulfillmentdate_Bulk();
                //2024-08-23 DO THIS IN SEPERATE DAILY AUTOMATE JOB NOW: "FULFILLMENTDATEUPDATES"
            }
            catch (Exception ex)
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine("ERROR :" + ex.ToString());
                }
                Returnvalue = false;
            }

            return Returnvalue;
        }

        public static Boolean Process_Product_Simple_Fulfillmentdate_Bulk(long Middlewareid = 0, int MaxRowstoProcess = 500)
        {
            Boolean Returnvalue = true;
            DataTable dt = new DataTable();
            String APIReturn = "";

            try
            {
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method FROM Communications..Middleware with (nolock) WHERE source_table = 'Herroom..ItemsFulfillmentdate' AND id = " + Middlewareid.ToString());
                }
                else
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT TOP " + MaxRowstoProcess.ToString() + " id, source_id, endpoint_method FROM Communications..Middleware with (nolock) WHERE source_table = 'Herroom..ItemsFulfillmentdate' AND Status = 100");
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
                            Console.WriteLine("ERROR for Fulfillmentdate Bulk Processing: " + dr["source_id"].ToString());
                        }
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, from_magento = '" + APIReturn + "' WHERE ID = " + dr["id"].ToString());
                    }
                    else //Should return new value
                    {
                        Console.WriteLine(APIReturn);
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = 'OK' WHERE ID = " + dr["id"].ToString() );
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Fulfillmentdate Bulk Processing(" + Middlewareid.ToString() + ") :: " + ex.ToString());
                Returnvalue = false;
            }

            return Returnvalue;
        }


        public static Boolean Process_Style_Price_Bulk(long Middlewareid = 0, int MaxRowstoProcess = 500)
        {
            Boolean Returnvalue = true;
            DataTable dt = new DataTable();
            String APIReturn = "";

            try
            {
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method FROM Communications..Middleware with (nolock) WHERE source_table = 'Herroom..StylePriceBulk' AND id = " + Middlewareid.ToString());
                }
                else
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT TOP " + MaxRowstoProcess.ToString() + " id, source_id, endpoint_method FROM Communications..Middleware with (nolock) WHERE source_table = 'Herroom..StylePriceBulk' AND Status = 100");
                }

                Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                {
                    APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");

                    if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                    {
                        //Middleware should already by set to 500
                        Console.WriteLine("ERROR for Price Style Bulk Processing: " + dr["source_id"].ToString());
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, from_magento = '" + APIReturn + "' WHERE ID = " + dr["id"].ToString());
                    }
                    else //Should return new value
                    {
                        Console.WriteLine(APIReturn);
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = 'OK' WHERE ID = " + dr["id"].ToString() + " AND STATUS NOT IN (700,701,702);");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Process_Item_Price(" + Middlewareid.ToString() + ") :: " + ex.ToString());
                Returnvalue = false;
            }

            return Returnvalue;
        }

        /// <summary>
        /// STYLENUMBER is a LIKE to "wac001%" will process all wacoal brand price updates, for instance
        /// </summary>
        /// <param name="Middlewareid"></param>
        /// <param name="Stylenumber"></param>
        /// <param name="MaxRowstoProcess"></param>
        /// <returns></returns>
        public static Boolean Process_Item_Price(long Middlewareid = 0, string Stylenumber = "", int MaxRowstoProcess = 400)
        {
            Boolean Returnvalue = true;
            DataTable dt = new DataTable();
            String APIReturn = "";
            int count = 1;

            try
            {
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method FROM Communications..Middleware with (nolock) WHERE source_table = 'Herroom..ItemPrice' AND id = " + Middlewareid.ToString());
                }
                else if (Stylenumber.Length > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT TOP " + MaxRowstoProcess.ToString() + " id, source_id, endpoint_method FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND source_table = 'Herroom..ItemPrice' AND batchid = '" + Stylenumber + "'");
                }
                else if (MaxRowstoProcess > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT TOP " + MaxRowstoProcess.ToString() + " id, source_id, endpoint_method FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND source_table = 'Herroom..ItemPrice' ORDER BY ID");
                }
                else
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND source_table = 'Herroom..ItemPrice' ORDER BY ID");
                }

                Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                {
                    APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");

                    if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                    {
                        //Middleware should already by set to 500
                        Console.WriteLine("ERROR for Process_Item_Price Processing: " + dr["source_id"].ToString());
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, from_magento = '" + APIReturn + "' WHERE ID = " + dr["id"].ToString());
                    }
                    else //Should return new value
                    {
                        if (MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine(APIReturn + " (" + count.ToString() + ")");
                        }
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = 'OK' WHERE ID = " + dr["id"].ToString() + " AND STATUS NOT IN (700,701,702);");
                    }
                    count++;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Process_Item_Price(" + Middlewareid.ToString() + ", " + Stylenumber + ") :: " + ex.ToString());
                Returnvalue = false;
            }

            return Returnvalue;
        }


        public static Boolean Product_Delete_Readd_Process()
        {
            Boolean bReturnvalue = true;
            DataTable dt = new DataTable();

            // Insert this row in HerTools 
            // Fetch Middleware rows 'DELETEREADD'
            // RUN Product_Delete_Readd(for source_id)

            try
            {
                dt = Helper.Sql_Misc_Fetch("SELECT id, source_id FROM communications..middleware WHERE Status=1000 AND endpoint_method='DELETEREADD' AND Source_table='herroom..styles' AND Posted <= Getdate(); ");
                Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                {
                    if (Product_Delete_Readd(dr["source_id"].ToString(), true, false, true))
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..middleware SET status=600 WHERE id = " + dr["id"].ToString());       
                    }
                    else
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..middleware SET status=700, tries=100 WHERE id = " + dr["id"].ToString());
                    }
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("ERROR: Product_Delete_Readd_Process() :: " + ex.ToString() );
                //Helper.Middleware_Status_Update(dt, 1000, "" );
                bReturnvalue = false;
            }


            return bReturnvalue;
        }

        public static Boolean Product_Delete_Readd(String Stylenumber, Boolean DeleteallSkus, Boolean DeleteInactiveOnly, Boolean ReAddActiveStyle)
        {
            Boolean bReturnvalue = true;
            DataTable dt = new DataTable();

            if (Generate_Product_Delete_Style_Rows(Stylenumber, DeleteallSkus, DeleteInactiveOnly))
            {
                if (Process_Product_Delete_Middleware())
                {
                    //ONLY RELOAD IF STYLE IS ACTIVE
                    if (ReAddActiveStyle)
                    {
                        dt = Helper.Sql_Misc_Fetch("SELECT ISNULL(Active,0) [Active] FROM Herroom..Styles WHERE Stylenumber = '" + Stylenumber + "' ;");

                        Console.WriteLine(dt.Rows[0]["active"].ToString());

                        if (dt.Rows[0]["active"].ToString() == "1" || dt.Rows[0]["active"].ToString().ToLower() == "true")
                        {
                            Helper.Sql_Misc_NonQuery("EXEC communications..[proc_mag_product_configurable_json_post_api] "
                                + " @Stylenumber = '" + Stylenumber + "' , @BlankProductLinks=1");
                        }
                    }
                }
            }

            return bReturnvalue;
        }

        public static Boolean Generate_Product_Delete_Style_Rows(String Stylenumber, Boolean DeleteallSkus, Boolean DeleteInactiveOnly)
        {
            Boolean bReturnvalue = true;
            String Sql;

            try
            {
                if (Stylenumber.Trim().Length > 0)
                {
                    Sql = "EXEC communications..[proc_mag_product_configurable_json_api_delete] @Stylenumber='" + Stylenumber + "' ";
                    if (DeleteallSkus)
                    {
                        Sql += ", @DeleteAllSkus=1";
                    }
                    else
                    {
                        Sql += ", @DeleteAllSkus=0";
                    }

                    if (DeleteInactiveOnly)
                    {
                        Sql += ", @InactiveONLY=1";
                    }
                    else
                    {
                        Sql += ", @InactiveONLY=0";
                    }

                    bReturnvalue = Helper.Sql_Misc_NonQuery(Sql);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Generate_Product_Delete_Style_Rows(" + Stylenumber + ") :: " + ex.ToString());
                bReturnvalue = false;
            }

            return bReturnvalue;
        }

        /// If Itemid > 0, Stylenumber is ignored by sproc
        public static Boolean Generate_Product_Delete_Item_Rows(String Stylenumber, long Itemid, Boolean DeleteInactiveOnly)
        {
            Boolean bReturnvalue = true;
            String Sql;

            try
            {
                Sql ="EXEC communications..[proc_mag_product_simple_json_api_delete] "
                    + "@Stylenumber ='" + Stylenumber.Trim() + "' "
                    + ", @Itemid=" + Itemid.ToString();

                if (DeleteInactiveOnly)
                {
                    Sql += ", @InactiveONLY=1";
                }
                else
                {
                    Sql += ", @InactiveONLY=0";
                }

                bReturnvalue = Helper.Sql_Misc_NonQuery(Sql);
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Generate_Product_Delete_Item_Rows(" + Stylenumber + ", " + Itemid.ToString() + " :: " + ex.ToString());
                bReturnvalue = false;
            }

            return bReturnvalue;
        }

        /// <summary>
        /// <param name="Middlewareid"></param>
        public static Boolean Process_Product_Delete_Middleware(long Middlewareid = 0)
        {
            Boolean bReturnvalue = true;
            DataTable dt = new DataTable();
            String APIReturn = "";
            try
            {
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, source_table, source_id FROM Communications..Middleware with (nolock) WHERE source_table IN ('herroom..styles','herroom..items') AND Endpoint_Method='DELETE' AND id = " + Middlewareid.ToString());
                }
                else
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT top 400 id, source_id, endpoint_method, source_table, source_id FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=1000 AND Endpoint_Method='DELETE' AND source_table IN ('herroom..styles','herroom..items') ORDER BY ID");
                }

                Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                {
                    APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");

                    if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                    {
                        //Middleware should already by set to 500
                        Console.WriteLine("ERROR for Process_Product_Configurable_Middleware Processing: " + dr["source_id"].ToString());
                    }
                    else if (APIReturn.ToLower().Contains("requested doesn't exist"))
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=601, from_magento = '" + APIReturn.Replace("'", "''") + "' WHERE ID = " + dr["id"].ToString());

                        if (dr["source_table"].ToString().ToLower() == "herroom..styles")
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Herroom..Styles SET Magentoid = NULL WHERE stylenumber = '" + dr["source_id"].ToString() + "' ");
                        }
                        else
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Herroom..Items SET Magentoid = NULL, MagentoSKU = NULL WHERE itemid = '" + dr["source_id"].ToString() + "' ");
                            Helper.Sql_Misc_NonQuery("DELETE communications..middleware WHERE source_table = 'herroom..itemslink' AND source_id = '" + dr["source_id"].ToString() + "'");
                        }
                    }
                    else //Should return 'TRUE' 
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + APIReturn.Replace("'", "''") + "' WHERE ID = " + dr["id"].ToString());

                        if (dr["source_table"].ToString().ToLower() == "herroom..styles")
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Herroom..Styles SET Magentoid = NULL WHERE stylenumber = '" + dr["source_id"].ToString() + "' ");
                        }
                        else
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Herroom..Items SET Magentoid = NULL, MagentoSKU = NULL WHERE itemid = '" + dr["source_id"].ToString() + "' ");
                            Helper.Sql_Misc_NonQuery("DELETE communications..middleware WHERE source_table = 'herroom..itemslink' AND source_id = '" + dr["source_id"].ToString() + "'"); 
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Process_Product_Delete_Middleware(" + Middlewareid.ToString() + ": " + ex.ToString());
                bReturnvalue = false;
                Helper.Middleware_Status_Update(dt, 100, "Status=200");
            }

            return bReturnvalue;
        }


        // NEED a batching system for this, should set status=200 immediately then 300 when actually processing somehow
        // for top 50 or whatever... 
        public static Boolean Process_Product_Configurable_Middleware(long Middlewareid = 0, int StatustoProcess = 100, string Envirionment = "")
        {
            Boolean bReturnvalue = true;
            DataTable dt = new DataTable();
            String APIReturn = "";
            long mid;

            try
            {
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,'') [batchid], to_magento FROM Communications..Middleware with (nolock) WHERE source_table = 'herroom..styles' AND id = " + Middlewareid.ToString());
                }
                else if (Middlewareid == -1)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT TOP 60 MM.id, source_id, endpoint_method, ISNULL(batchid, '') [batchid] "
                    + " FROM Communications..Middleware MM with (nolock) "
                    + " INNER JOIN communications..tempstyle SS on SS.Stylenumber = mm.source_id AND SS.Done = 12 "
                    + " WHERE status = 100 AND source_table = 'herroom..styles' ");

                }
                else if (Middlewareid == -2)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT TOP 200 MM.id, source_id, endpoint_method, ISNULL(batchid, '') [batchid] "
                     + " FROM Communications..Middleware MM with (nolock) "
                     + " WHERE status = 100 AND source_table = 'herroom..styles' AND Endpoint_Method = 'PUT' ");
                }
                else if (Middlewareid == -3)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=" + StatustoProcess.ToString() + " AND ISNULL(Batchnum,0) = 0 AND source_table = 'herroom..styles' AND Endpoint_method = 'POST' ORDER BY Posted, ID");
                }
                else if (Middlewareid < -4)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT top " + (-Middlewareid).ToString() + " id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=" + StatustoProcess.ToString() + " AND ISNULL(Batchnum,0) = 0 AND source_table = 'herroom..styles' ORDER BY Posted, ID");
                }
                else
                {
                    // RUN SPROC TO FETCH BatchIDS
                    dt = Helper.Sql_Misc_Fetch("SELECT top 800 id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=" + StatustoProcess.ToString() + " AND ISNULL(Batchnum,0) = 0 AND source_table = 'herroom..styles' ORDER BY Posted, ID");
                }

                Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                { 

                    if (MagetnoProductAPI.DevMode > 1 && Middlewareid > 0)
                    {
                        Console.WriteLine(dr["to_magento"].ToString());
                    }
                    APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString()), Envirionment).Replace("\"", "");

                    if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                    {
                        //Middleware should already by set to 500
                        Console.WriteLine("ERROR for Process_Product_Configurable_Middleware Processing: " + dr["source_id"].ToString());
                    }
                    else if (APIReturn.Length > 30 && APIReturn.ToLower().Contains("url key for specified store"))
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, error_message = 'URL key for specified store already exists' WHERE ID = " + Middlewareid.ToString());
                    }
                    else if (APIReturn.ToLower().Contains("message") && (APIReturn.ToLower().Contains("try again") || APIReturn.ToLower().Contains("request does not match any route")))
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, tries=10, error_message = '" + APIReturn.Replace("'", "''") + "' WHERE ID = " + dr["id"].ToString());
                    }
                    else if (APIReturn.ToLower().Contains("message") && (APIReturn.ToLower().Contains("have the same set of attribute values") || APIReturn.ToLower().Contains("try again")))
                    {
                        //redo json 
                        if (dr["endpoint_method"].ToString().ToUpper() == "POST")
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=100, error_message = '" + APIReturn.Replace("'", "''") + "' WHERE ID = " + dr["id"].ToString());
                            Helper.Sql_Misc_NonQuery("EXEC communications..[proc_mag_product_configurable_json_post_api] @stylenumber='" + dr["source_id"].ToString() + "', @BlankProductLinks=1, @MiddlwareID =" + dr["id"].ToString() + ";");
                        }
                        else
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, tries=10, error_message = '" + APIReturn.Replace("'", "''") + "' WHERE ID = " + dr["id"].ToString());
                            Helper.Sql_Misc_NonQuery("EXEC communications..[proc_mag_product_configurable_json_api] @Stylenumber='" + dr["source_id"].ToString() + "', @BlankProductLinks=1;");
                        }
                    }
                    else //Should return new value
                    {
                        if (long.TryParse(APIReturn, out mid) && dr["endpoint_method"].ToString().ToUpper() == "POST")
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + mid.ToString().Replace("'", "''") + "' WHERE ID = " + dr["id"].ToString() + " AND Status NOT IN (700,701,702);");
                            Helper.Sql_Misc_NonQuery("UPDATE Herroom..Styles SET Magentoid = " + mid.ToString() + " WHERE stylenumber = '" + dr["source_id"].ToString() + "' AND Magentoid IS NULL");
                        }
                        else if (APIReturn.ToLower() == "{message:url key for specified store already exists.}")
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, error_message = '" + APIReturn.Replace("'", "''") + "' WHERE ID = " + dr["id"].ToString());
                        }
                        else
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + APIReturn.Replace("'", "''") + "' WHERE ID = " + dr["id"].ToString() + " AND Status NOT IN (700,701,702)");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Process_Product_Configurable_Middleware(" + Middlewareid.ToString() + ": " + ex.ToString());
                bReturnvalue = false;
                Helper.Middleware_Status_Update(dt, 100, "Status=200");
            }

            return bReturnvalue;
        }

        /// <summary>
        /// This call sproc to rewrite the JSON using info that was missing before for brand, colors, sizes, collections, OST, and anything else
        /// </summary>
        /// <param name="Middlewareid"></param>
        /// <returns></returns>
        public static Boolean ReProcess_Product_Configurable_Middleware(long Middlewareid = 0)
        {
            Boolean bReturnvalue = true;
            DataTable dt;

            if (Middlewareid > 0)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT Source_id, id FROM communications..middleware with (nolock) WHERE Posted <= Getdate() AND status=100 AND batchnum=100 AND source_table = 'herroom..styles' AND id = " + Middlewareid.ToString());
            }
            else  // DO ALL STYLES STATUS=100 AND BatchID = @Sourceid (Stylenumber)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT Source_id, id FROM communications..middleware with (nolock) WHERE Posted <= Getdate() AND status=100 AND batchnum=100 AND source_table = 'herroom..styles' ORDER BY ID");
            }

            foreach (DataRow dr in dt.Rows)
            {
                //redo json 
                Helper.Sql_Misc_NonQuery("EXEC communications..[proc_mag_product_configurable_json_post_api] @stylenumber = '" + dr["source_id"].ToString() + "', @MiddlwareID =" + dr["id"].ToString() + ", @BlankProductLinks=1;");
            }

            return bReturnvalue;
        }

        /// <summary>
        /// This call sproc to rewrite the JSON using info that was missing before for brand, colors, sizes, collections, OST, and anything else
        /// </summary>
        /// <param name="Middlewareid"></param>
        /// <returns></returns>
        public static Boolean ReProcess_Product_Simple_Middleware(long Middlewareid = 0)
        {
            Boolean bReturnvalue = true;
            DataTable dt;

            if (Middlewareid > 0)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT Source_id, id, ISNULL(batchid,'') [batchid] FROM communications..middleware with (nolock) WHERE Posted <= Getdate() AND source_table = 'herroom..items' AND id = " + Middlewareid.ToString());
            }
            else  // DO ALL STYLES STATUS=100 AND BatchID = @Sourceid (Stylenumber)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT Source_id, id, ISNULL(batchid,'') [batchid] FROM communications..middleware with (nolock) WHERE Posted <= Getdate() AND status=100 AND batchnum=100 AND source_table = 'herroom..items' ORDER BY ID");
            }

            foreach (DataRow dr in dt.Rows)
            {
                Console.WriteLine("EXEC communications..[proc_mag_product_simple_json_api] @itemid = " + dr["source_id"].ToString() + ", @upc='', @MiddlewareID =" + dr["id"].ToString());
                //redo json and update batchnum=0 (sproc does all that is needed)
                Helper.Sql_Misc_NonQuery("EXEC communications..[proc_mag_product_simple_json_api] @itemid = " + dr["source_id"].ToString() + ", @upc='', @MiddlewareID =" + dr["id"].ToString());
             }

            return bReturnvalue;
        }

        /// <summary>
        /// This call sproc to rewrite the JSON using info that was missing before for SKUs
        /// </summary>
        /// <param name="Middlewareid"></param>
        /// <returns></returns>
        public static Boolean ReProcess_Product_Links_Middleware(long Middlewareid = 0)  // MiddlewareInsert_Product_Links
        {
            Boolean bReturnvalue = true;
            DataTable dt;

            //FIRST, PUSH FORWARD ANY Links FOR STYLES WHERE Magentoid IS NULL
            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET Posted = Dateadd(minute, 15, Getdate()) " +
                " FROM Communications..Middleware MM with(nolock) " +
                " INNER JOIN Herroom..Styles RSS ON RSS.StyleNumber = MM.source_id AND ISNULL(RSS.magentoid, 0) = 0 " +
                " WHERE Posted <= Getdate() AND Status = 100 " +
                " AND ISNULL(Batchnum, 0) = 100 AND source_table = 'herroom..itemslink' ");

            if (Middlewareid > 0)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT Source_id, id FROM communications..middleware with (nolock) WHERE Posted <= Getdate() AND status=100 AND batchnum=100 AND source_table = 'herroom..itemslink' AND id = " + Middlewareid.ToString());
            }
            else  // DO ALL STYLES STATUS=100 AND BatchID = @Sourceid (Stylenumber)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT Source_id, id FROM communications..middleware with (nolock) WHERE Posted <= Getdate() AND status=100 AND batchnum=100 AND source_table = 'herroom..itemslink' ORDER BY POSTED, ID");
            }

            foreach (DataRow dr in dt.Rows)
            {
                //redo json 
                Helper.Sql_Misc_NonQuery("EXEC communications..[proc_mag_product_links_json_api_child] @Stylenumber = '" + dr["Source_id"].ToString() + "', @Middlewareid = " + dr["id"].ToString());
            }

            return bReturnvalue;
        }


        public static Boolean Process_Product_InStock_Middleware(long Middlewareid = 0, int Status = 100)   
        {
            Boolean bReturnvalue = true;
            DataTable dt;
            String APIReturn;
            long mid;

           if (Middlewareid > 0)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT Source_id, id FROM communications..middleware with (nolock) WHERE source_table = 'herroom..StylesInStock' AND id = " + Middlewareid.ToString());
            }
            else  // DO ALL STYLES STATUS=100 AND BatchID = @Sourceid (Stylenumber)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT Source_id, id FROM communications..middleware MM with(nolock) WHERE Posted <= Getdate() AND status = " + Status.ToString() + " AND source_table = 'herroom..StylesInStock' ORDER BY POSTED, ID");
            }

            Helper.Middleware_Status_Update(dt, 200);

            foreach (DataRow dr in dt.Rows)
            {
    
                APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");
                Console.WriteLine(APIReturn);
                Console.WriteLine("-------------------");

                if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                {
                    Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, error_message = '" + APIReturn + "' WHERE ID = " + dr["id"].ToString());
                    Console.WriteLine("ERROR for Process_Product_Simple_Middleware Processing: " + dr["id"].ToString());
                }
                else if (APIReturn.ToLower().Contains("request does not match any route"))
                {
                    Console.WriteLine("@!!!");
                    Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, error_message = '" + APIReturn + "' WHERE ID = " + dr["id"].ToString());

                }
                else
                {
                    if (long.TryParse(APIReturn, out mid))
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + mid.ToString() + "' WHERE ID = " + dr["id"].ToString());
                    }
                }
            }

            return bReturnvalue;
        }


        public static Boolean Process_Product_Simple_Middleware(long Middlewareid = 0, int StatustoProcess = 100, int MaxRowstoProcess = 200)
        {
            Boolean bReturnvalue = true;
            DataTable dt = new DataTable();
            String APIReturn = "";
            long mid;
            string Itemid;
            int stringHelper;
            string tmp;

            try
            { 
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,'') [batchid], ISNULL(module,'') [module] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND source_table = 'herroom..items' AND id = " + Middlewareid.ToString());
                }
                else if (Middlewareid == -1)
                {

                    dt = Helper.Sql_Misc_Fetch("SELECT TOP 500 MM.id, source_id, endpoint_method, ISNULL(batchid,'') [batchid], ISNULL(module,'') [module] FROM Communications..Middleware MM with (nolock) "
                        + " INNER JOIN TempsKU SKU on convert(varchar(32), itemid) = MM.source_id "
                        + " WHERE source_table = 'herroom..items' "
                        + " AND SKU.done = 11 AND MM.status = 101");
                }
                else if (Middlewareid == -3)
                {
                    //ONLY RUNS POST (new) SKUS
                    dt = Helper.Sql_Misc_Fetch("SELECT TOP " + MaxRowstoProcess.ToString() + " id, source_id, endpoint_method, ISNULL(batchid,'') [batchid], ISNULL(module,'') [module] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=" + StatustoProcess.ToString() + " AND ISNULL(Batchnum,0) = 0 AND source_table = 'herroom..items' AND Endpoint_Method='POST' AND Len(Ltrim(RTrim(ISNULL(to_magento,'')))) > 0 ORDER BY endpoint_Method, ID");
                }
                else
                {
                    // 2 SELECTS: 1) Get all of the SKUS in Batchids under styles in status 600 
                    //  2) GET 200 all non-batch related IDS (PUTS usually) 

                    dt = Helper.Sql_Misc_Fetch("SELECT TOP " + MaxRowstoProcess.ToString() + " id, source_id, endpoint_method, ISNULL(batchid,'') [batchid], ISNULL(module,'') [module] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=" + StatustoProcess.ToString() + " AND ISNULL(Batchnum,0) = 0 AND source_table = 'herroom..items' AND Len(Ltrim(RTrim(ISNULL(to_magento,'')))) > 0 ORDER BY endpoint_Method, ID");
                }

                Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                {
                    APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");

                    if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                    {
                        //Middleware should already by set to 500
                        Console.WriteLine("ERROR for Process_Product_Simple_Middleware Processing: " + dr["id"].ToString());
                    }
                    else if (APIReturn.Length > 10 && APIReturn.Substring(0, 8).ToLower() ==  "{message" && (APIReturn.Contains("try again") || APIReturn.ToLower().Contains("request does not match any route")))
                    {
                        tmp = APIReturn.Replace("'", "");
                      
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, error_message = '" + tmp + "' WHERE ID = " + dr["id"].ToString());
                    }
                    else if (APIReturn.Length >= 4 && APIReturn.Substring(0, 4) == "{id:")
                    {
                        stringHelper = APIReturn.IndexOf("sku");
                        Itemid = APIReturn.Substring(4, stringHelper - 4);
                        Itemid = Itemid.Replace(",", "");
   
                        if (long.TryParse(Itemid, out mid))
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + mid.ToString() + "' WHERE ID = " + dr["id"].ToString());
                            Helper.Sql_Misc_NonQuery("UPDATE Herroom..Items SET Magentoid = " + mid.ToString() + " WHERE itemid = '" + dr["source_id"].ToString() + "' AND Magentoid IS NULL");
                        }
                    }
                    else //Should return new value
                    {
                        if (long.TryParse(APIReturn, out mid) && dr["endpoint_method"].ToString().ToUpper() == "POST")
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + mid.ToString() + "' WHERE ID = " + dr["id"].ToString());
                            Helper.Sql_Misc_NonQuery("UPDATE Herroom..Items SET Magentoid = " + mid.ToString() + ", MagentoSKU = '" + dr["module"].ToString() + "' WHERE itemid = '" + dr["source_id"].ToString() + "' AND Magentoid IS NULL");           
                        }
                        else if (APIReturn.ToLower() == "{message:url key for specified store already exists.}")
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, error_message = '" + APIReturn.Replace("'", "''") + "' WHERE ID = " + dr["id"].ToString());
                        }
                        else
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + APIReturn.Replace("'", "''") + "' WHERE ID = " + dr["id"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Process_Product_Simple_Middleware(" + Middlewareid.ToString() + ": " + ex.ToString());

                bReturnvalue = false;
                Helper.Middleware_Status_Update(dt, 100, "Status=200");
            }

            return bReturnvalue;
        }

        /// This is no longer needed to be set
        /// it was set ONLY on POST of a config due to the way the websiteids were arranged
        /// but that has been updated and this isn't needed
        public static Boolean Process_Product_Visibility_Middleware(long Middlewareid = 0)
        {
            Boolean bReturnvalue = true;
            DataTable dt;
            String APIReturn = "";
            long mid;
            //FIRST, PUSH FORWARD ANY OPTION FOR STYLES WHERE Magentoid IS NULL
            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET Posted = Dateadd(minute, 20, Getdate()) " +
                " FROM Communications..Middleware MM with(nolock) " +
                " INNER JOIN Herroom..Styles RSS ON RSS.StyleNumber = MM.source_id AND ISNULL(RSS.magentoid, 0) = 0 " +
                " WHERE Posted <= Getdate() AND Status = 100 " +
                " AND source_table IN ('herroom..styleVisibility','herroom..itemVisibility') ");

            try
            {
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND source_table IN ('herroom..styleVisibility', 'herroom..itemVisibility') AND ISNULL(Batchnum,0) = 0 AND id = " + Middlewareid.ToString());
                }
                else
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND ISNULL(Batchnum,0) = 0 AND source_table IN ('herroom..styleVisibility', 'herroom..itemVisibility') ORDER BY ID");
                }

                Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                {
                    APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");

                    if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                    {
                        //Middleware should already by set to 500
                        Console.WriteLine("ERROR for Process_Product_Simple_Middleware Processing: " + dr["source_id"].ToString());
                    }
                    else //Should return new value
                    {
                        if (long.TryParse(APIReturn, out mid))
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + mid.ToString() + "' WHERE ID = " + dr["id"].ToString());
                        }
                        else
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = 'DONE' WHERE ID = " + dr["id"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Process_Product_Visibility_Middleware(" + Middlewareid.ToString() + ": " + ex.ToString());
                bReturnvalue = false;
            }

            return bReturnvalue;
        }

        public static Boolean Process_Product_Links_Middleware(string Stylenumber = "" , int MaxRowstoProcess = 0)
        {
            Boolean bReturnvalue = true;
            DataTable dt;
            String APIReturn = "";

            if (Stylenumber == "")
            {
                if (MaxRowstoProcess > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT TOP " + MaxRowstoProcess.ToString() + " id, source_id, endpoint_method FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND ISNULL(Batchnum,0) = 0 AND source_table = 'herroom..itemslink' ORDER BY POSTED;");
                }
                else
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT TOP 300 id, source_id, endpoint_method FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND ISNULL(Batchnum,0) = 0 AND source_table = 'herroom..itemslink' ORDER BY POSTED;");
                }
            }
            else if (Stylenumber == "-1")
            {
                dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method FROM Communications..Middleware with (nolock) WHERE Status=102 AND source_table = 'herroom..itemslink' ORDER BY POSTED, ID");
            }
            else if (Stylenumber == "-2")  
            {
                dt = Helper.Sql_Misc_Fetch("SELECT TOP 100 id, source_id, endpoint_method FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=106 AND ISNULL(Batchnum,0) = 0 AND source_table = 'herroom..itemslink' ORDER BY TEST, Source_id ;");
            }
            else
            {
                if (MaxRowstoProcess > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT TOP " + MaxRowstoProcess.ToString() + " id, source_id, endpoint_method FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND ISNULL(Batchnum,0) = 0 AND source_table = 'herroom..itemslink' " +
                         " AND source_id IN (SELECT Convert(varchar(32), itemid) FROM Herroom..Items WHERE stylenumber = '" + Stylenumber + "') ORDER BY POSTED, ID");
                }
                else
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND ISNULL(Batchnum,0) = 0 AND source_table = 'herroom..itemslink' " +
                        " AND source_id IN (SELECT Convert(varchar(32), itemid) FROM Herroom..Items WHERE stylenumber = '" + Stylenumber + "') ORDER BY POSTED, ID");
                }
            }
            Helper.Middleware_Status_Update(dt, 200);

            Stopwatch sw = new Stopwatch();
            sw.Start();
            foreach (DataRow dr in dt.Rows)
            {
                APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");

                Console.WriteLine(APIReturn);

                if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                {
                    //Middleware should already by set to 500
                    APIReturn = APIReturn.Replace("'", "");
                    Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=701, error_message = '" + APIReturn + "' WHERE ID = " + dr["id"].ToString());
                    Console.WriteLine("ERROR for Process_Product_Simple_Middleware Processing: " + dr["source_id"].ToString());
                    bReturnvalue = false;
                }
                else if (APIReturn.Length > 4 && APIReturn.Contains("doesn't exist"))
                {
                    APIReturn = APIReturn.Replace("'", "");
                    Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=701, error_message = '" + APIReturn + "' WHERE ID = " + dr["id"].ToString());
                }
                else //Should return new value
                {
                    APIReturn = APIReturn.Replace("'", "");

                    //Reprocess... seeing these even though the Style options have run sucessfully
                    if (APIReturn.ToLower().Contains("the parent product doesnt have configurable product options"))
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, Posted = Dateadd(minute, 20, Posted), Tries = (ISNULL(tries,0) + 1), from_magento = '" + APIReturn + "' WHERE ID = " + dr["id"].ToString() + " AND Tries >= 6");

                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=100, Posted = Dateadd(minute, 20, Posted), Tries = (ISNULL(tries,0) + 1), from_magento = '" + APIReturn + "' WHERE ID = " + dr["id"].ToString() + " AND Tries < 6");
                    }
                    else
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + APIReturn + "' WHERE ID = " + dr["id"].ToString());
                    }
                }

                //1000 * 18:00 * 60 seconds = 1,080,000
                if (sw.ElapsedMilliseconds > 1080000)
                {
                    break;
                }
            }

            //Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=101 WHERE status=200 AND source_table = 'herroom..itemslink'; ");
            Helper.Middleware_Status_Update(dt, 100, " status = 200 ");

            return bReturnvalue;
        }


        // FOR NOW, CALL THIS FOR EACH STYLENUMBER GOING LIVE... AFTER SKUS ARE LOADED
        public static Boolean Create_Product_Links_Middleware(string Stylenumber, int RecreateLinks = 0)
        {
            Boolean bReturnvalue = true;

            if (RecreateLinks != 0)
            {
                RecreateLinks = 1;
            }

            try
            {
                //GENERATE NEW LINKS (NOT ALREADY PROCESSED IN MIDDLEWARE NOT IN STATUS (100, 200, 500, 600, 700, 800)
                if (Helper.Sql_Misc_NonQuery("EXEC Communications..[proc_mag_product_links_json_api_child] @Stylenumber='" + Stylenumber + "', @Recreate=" + RecreateLinks.ToString() + ";"))
                {
                    Process_Product_Links_Middleware(Stylenumber);
                }
            }
            catch (Exception ex)
            { 
                Console.WriteLine("ERROR: Process_Product_Links_Middleware(" + Stylenumber + ": " + ex.ToString());
                bReturnvalue = false;
            }

            return bReturnvalue;
        }

        //OVERLOAD
        public static Boolean Generate_Product_Links_Middleware(String Stylenumber = "", int RecreateLinks = 0)
        {
            Boolean bReturnvalue = true;
            DataTable dtStyles = new DataTable();

            try
            {
                if (Stylenumber.Length > 0)
                {
                    dtStyles = Helper.Sql_Misc_Fetch("SELECT ISNULL(source_id,'') [source_id], id " +
                       " FROM Communications..Middleware MM " +
                       " WHERE Status = 100 AND Posted <= Getdate() AND Source_Table = 'CreateItemLinks' " +
                       " AND Source_id = '" + Stylenumber + "'; ");
                }
                else
                {
                    //CLEAR OUT QUEUE STYLES WITH NO ACTIVE Items
                    Helper.Sql_Misc_NonQuery("UPDATE Middleware "
                        + " SET Tries = ISNULL(Tries, 0) + 1, Posted = Dateadd(day, 1, Posted) "
                        + " FROM Middleware MM WHERE Status = 100 AND Posted <= Getdate() AND source_table = 'createitemlinks' AND TRIES < 6 "
                        + " AND 0 = (SELECT count(*) FROM herroom..items WHERE convert(varchar(20), itemid) = source_id AND magentoid IS NOT NULL AND Active = 1)");

                    //REQUEUE STYLES WITH NO ACTIVE Items
                    Helper.Sql_Misc_NonQuery("UPDATE Middleware "
                        + " SET Tries = 10, Status = 800 "
                        + " FROM Middleware MM WHERE Status = 100 AND Posted <= Getdate() AND source_table = 'createitemlinks' AND TRIES >= 6 "
                        + " AND 0 = (SELECT count(*) FROM herroom..items WHERE convert(varchar(20), itemid) = source_id AND magentoid IS NOT NULL AND Active = 1)");

                    //CHECK FOR NEW SKUS TO BE INSERTED FIRST
                    dtStyles = Helper.Sql_Misc_Fetch("SELECT ISNULL(source_id,'') [source_id], id " +
                        "FROM Communications..Middleware MM " +
                        "WHERE Status = 100 AND Posted <= Getdate() AND Source_Table = 'CreateItemLinks' " +
                        "AND 0 = (SELECT COUNT(*) FROM Communications..Middleware MX WHERE Status = 100 AND source_table = 'Herroom..Items' AND endpoint_method = 'POST' AND MX.batchid=MM.batchid) ;");
                }

                Helper.Middleware_Status_Update(dtStyles, 200);

                foreach (DataRow drStyle in dtStyles.Rows)
                {
                    if (Create_Product_Links_Middleware(drStyle["Source_id"].ToString(), RecreateLinks))
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET Status=600 WHERE id = " + drStyle["id"].ToString());
                    }
                    else
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET Status=700, error_message='Unknown Issue, See herroom..ItemsLink rows for Stylenumber (batchid)' WHERE id = " + drStyle["id"].ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Process_Product_Links_Middleware() :: " + ex.ToString());
                bReturnvalue = false;
                Helper.Middleware_Status_Update(dtStyles, 100, "Status=200");
            }

            return bReturnvalue;
        }

        public static Boolean Process_Product_Options_Middleware(long Middlewareid = 0, string Stylenumber = "")
        {
            Boolean bReturnvalue = true;
            DataTable dt = new DataTable();
            String APIReturn = "";
            long mid;

            try
            {
                //FIRST, PUSH FORWARD ANY OPTION FOR STYLES WHERE Magentoid IS NULL
                Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET Posted = Dateadd(minute, 10, Getdate()) " +
                    " FROM Communications..Middleware MM with(nolock) " +
                    " INNER JOIN Herroom..Styles RSS ON RSS.StyleNumber = MM.source_id AND ISNULL(RSS.magentoid, 0) = 0 " +
                    " WHERE Posted <= Getdate() AND Status = 100 " +
                    " AND source_table = 'herroom..styleOptions'  ");
                //    " AND ISNULL(Batchnum, 0) = 0 AND source_table = 'herroom..styleOptions'  ");

                if (Middlewareid > 0)
                {
                    //LEAVE THIS SELECT ALONE FOR NOW - FOR TESTING 
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method FROM Communications..Middleware with (nolock) WHERE source_table = 'herroom..styleOptions' AND id = " + Middlewareid.ToString());

                    //dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method  " +
                    //   " FROM Communications..Middleware MM with(nolock) " +
                    //   " INNER JOIN Herroom..Styles RSS ON RSS.StyleNumber = MM.source_id AND ISNULL(RSS.magentoid, 0) > 0 " +
                    //   " WHERE Posted <= Getdate() AND Status = 100 " +
                    //   " AND ISNULL(Batchnum, 0) = 0 AND source_table = 'herroom..styleOptions' AND ISNULL(Batchnum,0) = 0 AND id = " + Middlewareid.ToString());
                }
                else if (Stylenumber.Length > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND source_table = 'herroom..styleOptions' AND Source_id = '" + Stylenumber + "'; ");
                }
                else
                {
                    //WITH THE UPDATE ADDED ABOVE, THIS MAY BE OVERCODED, REVISIT LATER
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method  " +
                        " FROM Communications..Middleware MM with(nolock) " +
                        " INNER JOIN Herroom..Styles RSS ON RSS.StyleNumber = MM.source_id AND ISNULL(RSS.magentoid, 0) > 0 " +
                        " WHERE Posted <= Getdate() AND Status = 100 " +
                        " AND ISNULL(Batchnum, 0) = 0 AND source_table = 'herroom..styleOptions' ORDER BY ID");            
                }

                Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                {
                    APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");

                    if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                    {
                        //Middleware should already by set to 500
                        Console.WriteLine("ERROR for Process_Product_Options_Middleware Processing: " + dr["source_id"].ToString());
                    }
                    else //Should return new value
                    {
                        if (long.TryParse(APIReturn, out mid))
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + mid.ToString() + "' WHERE ID = " + dr["id"].ToString());
                        }
                        else
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = 'DONE' WHERE ID = " + dr["id"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Process_Product_Links_Middleware(" + Middlewareid.ToString() + ": " + ex.ToString());
                bReturnvalue = false;
                Helper.Middleware_Status_Update(dt, 100, "Status=200");
            }

            return bReturnvalue;
        }


        /// <summary>
        /// ItemID > Stylenumber, sproc will only do inserts for SKU if itemid passed in > 0
        /// FOR TESTING
        /// </summary>
        /// <param name="Stylenumber"></param>
        /// <param name="Itemid"></param>
        /// <returns></returns>
        public static Boolean MiddlewareInsert_Product_Visibility_Insert(String Preprocessorid, String Stylenumber, String Itemid = "0")
        {
            Boolean bReturnvalue = true;
            //Does Middleware Inserts based on Styles/Items Gender 
            Helper.Sql_Misc_NonQuery("EXEC Communications..[proc_mag_product_visibility_json_api] '" + Stylenumber + "', " + Itemid + ", '" + Preprocessorid + "'");

            return bReturnvalue;
        }

        //FOR TESTING
        public static Boolean MiddlewareInsert_Product_Links_Insert(String Stylenumber, String Middlewareid = "0")
        {
            Boolean bReturnvalue = true;

            Helper.Sql_Misc_NonQuery("EXEC Communications..[proc_mag_product_links_json_api] '" + Stylenumber + "', " + Middlewareid);

            return bReturnvalue;
        }

        /// <summary>
        /// ItemID > Stylenumber, sproc will only do inserts for SKU if itemid passed in > 0
        /// FOR TESTING
        /// </summary>
        /// <param name="Stylenumber"></param>
        /// <param name="Itemid"></param>
        /// <returns></returns>
        public static Boolean MiddlewareInsert_Product_Options_Insert(String Stylenumber)
        {
            Boolean bReturnvalue = true;

            Helper.Sql_Misc_NonQuery("EXEC Communications..[proc_mag_product_options_json_api] '" + Stylenumber + "';"); 

            return bReturnvalue;
        }

        public static Boolean Middleware_Get(long Middlewareid = 0, string SourceTable = "")
        {
            Boolean bReturnvalue = true;
            DataTable dt;
            String APIReturn = "";
            String Sql;

            try
            {
                if (Middlewareid > 0)
                {
                    Sql = "SELECT id FROM communications..Middleware WHERE Endpoint_Method = 'GET' AND ID = " + Middlewareid.ToString();
                }
                else
                {
                    Sql = "SELECT id FROM communications..Middleware WHERE Status = 100 AND endpoint_method = 'GET' ";
                    if (SourceTable.Length > 0)
                    {
                        Sql += " AND Source_Table = '" + SourceTable + "' ";
                    }

                    Sql += " ORDER BY ID DESC;";
                }
      
                dt = Helper.Sql_Misc_Fetch(Sql);

                Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                {
                    //Helper.MagentoApiPush() handles all of this for GETs
                    APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Middleware_Get(" + Middlewareid.ToString() + "): " + ex.ToString());
                bReturnvalue = false;
            }

            return bReturnvalue;
        }


        public static Boolean StockItems_Adjustment(long Middlewareid = 0)
        {
            Boolean bReturnvalue = true;
            DataTable dt;
            String APIReturn = "";

            //THE ROWS ARE CREATED IN SPROC proc_mag_stockitems_0_audit_api which is called from Sql Job: Magento 0 Inventory Sync  
            if (Middlewareid > 0)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT id, source_id FROM Communications..Middleware WHERE endpoint_method = 'PUT' AND source_table = 'herroom..stockitems' AND id = " + Middlewareid.ToString());
            }
            else
            {
                Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET STATUS = 700 FROM Communications..Middleware WHERE Status = 100 AND endpoint_method = 'PUT' AND source_table = 'herroom..stockitems' AND len(isnull(to_magento, '')) = 0;");

                dt = Helper.Sql_Misc_Fetch("SELECT top 2000 id, source_id FROM Communications..Middleware WHERE Status=100 AND endpoint_method = 'PUT' AND source_table = 'herroom..stockitems' AND len(isnull(to_magento,'')) > 0;");
            }

            Helper.Middleware_Status_Update(dt, 200);

            foreach (DataRow dr in dt.Rows)
            {
                APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");

                if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                {
                    //Middleware should already by set to 500
                    Console.WriteLine("ERROR StockItems_Adjustment: " + dr["source_id"].ToString());
                    Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, from_magento = 'API ERROR... " + APIReturn + "' WHERE ID = " + dr["id"].ToString());
                }
                else //Should return new value
                {
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + APIReturn + "' WHERE ID = " + dr["id"].ToString());
                    }
                }
            }

            return bReturnvalue;
        }

        public static Boolean StockItems_Adjustment_Get(long Middlewareid = 0)
        {
            Boolean bReturnvalue = true;
            DataTable dt;
            String APIReturn = "";

            //THE ROWS ARE CREATED IN SPROC proc_mag_stockitems_0_audit_api which is called from Sql Job: Magento 0 Inventory Sync  
            if (Middlewareid > 0)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_name FROM Communications..Middleware WHERE endpoint_method = 'GET' AND source_table = 'herroom..stockitems' AND id = " + Middlewareid.ToString());
            }
            else
            {
                dt = Helper.Sql_Misc_Fetch("SELECT top 1000 id, source_id, endpoint_name FROM Communications..Middleware WHERE Status=100 AND endpoint_method = 'GET' AND source_table = 'herroom..stockitems'");
            }

            Helper.Middleware_Status_Update(dt, 200);

            foreach (DataRow dr in dt.Rows)
            {
                APIReturn = Helper.MagentoApiGet(dr["endpoint_name"].ToString(), long.Parse(dr["id"].ToString())).Replace("\"", "");

                if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                {
                    //Middleware should already by set to 500
                    Console.WriteLine("ERROR StockItems_Adjustment: " + dr["source_id"].ToString());
                    Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, from_magento = 'API ERROR... " + APIReturn + "' WHERE ID = " + dr["id"].ToString());
                }
                else //Should return new value
                {
                    {
                       // Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + APIReturn + "' WHERE ID = " + dr["id"].ToString());
                    }
                }
            }

            return bReturnvalue;
        }


        public static Boolean Product_Review_Comment_Check(int CustCommentsid = 0)
        {
            Boolean bReturnvalue = true;
            DataTable dt;
            String APIGet = "";
            String Status;
            dynamic dJson;

            if (CustCommentsid > 0)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT Magentoid, review, vetted FROM Hercust..CustComments WHERE id = " + CustCommentsid.ToString());
            }
            else
            {
                dt = Helper.Sql_Misc_Fetch("SELECT  Magentoid, review, vetted FROM Hercust..CustComments WHERE ISNULL(Magentoid,0) > 0 AND Vetted = 1 AND MagentoVetted2 IS NULL Order BY id");
            }

            foreach (DataRow dr in dt.Rows)
            {
                APIGet = Helper.MagentoApiGet("V1/reviews/" + dr["magentoid"].ToString());

                if (MagetnoProductAPI.DevMode > 1)
                {
                    Console.WriteLine(APIGet);
                }

                if (APIGet.Length > 4 && APIGet.Substring(0, 5).ToUpper() == "ERROR")
                {
                    Helper.Sql_Misc_NonQuery("UPDATE Hercust..CustComments SET MagentoVetted = -1 WHERE Magentoid = " + dr["magentoid"].ToString());
                }
                else
                {
                    dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(APIGet);

                    if (dJson.review_status != null)
                    {
                        try
                        {
                            Status = dJson.review_status;

                            if (MagetnoProductAPI.DevMode > 1)
                            {
                                Console.WriteLine("Status: " + Status.ToString());
                            }

                            if (dr["vetted"].ToString() == "1")
                            {
                                dJson.review_status = 1;        //APPROVED
                            }
                            else
                            {
                                dJson.review_status = 3;      //NOT APPROVED
                                //dJson.review_status = 2;      //PENDING
                            }

                            if (dJson.detail != dr["review"].ToString())
                            {
                                Console.WriteLine(dJson.detail);
                                Console.WriteLine(" ----- ");
                                Console.WriteLine(dr["review"].ToString());
                                Console.WriteLine(" ----------------------------------------- ");
                                Helper.Sql_Misc_NonQuery("UPDATE hercust..custcomments SET magentovetted2 = 1 WHERE Magentoid = " + dr["magentoid"].ToString());
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("ERROR: NO Status RETURNED (1): " + ex.ToString());
                        }
                    }
                }
            }

            return bReturnvalue;
        }
        /// <summary>
        /// calls Product_Review_Status_Update(CustCommentsid) to do processing
        /// </summary>
        /// <returns></returns>
        public static Boolean Process_Review_Updates_Middleware()
        {
            Boolean Returnvalue = true;
            DataTable dt;

            dt = Helper.Sql_Misc_Fetch("SELECT ID, source_id FROM communications..middleware WHERE Status = 100 AND source_table = 'hercust..custcomments';");

            Helper.Middleware_Status_Update(dt, 200);
            foreach(DataRow dr in dt.Rows)
            {
                try
                {
                    if (Product_Review_Status_Update(int.Parse(dr["source_id"].ToString())))
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE communications..middleware SET Status = 600 WHERE id = " + dr["id"].ToString());
                    }
                    else
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE communications..middleware SET Status = 700, error_message='Product_Review_Status_Update error' WHERE id = " + dr["id"].ToString());
                    }
                }
                catch(Exception ex)
                {
                    if (MagetnoProductAPI.DevMode >0)
                    {
                        Console.WriteLine("ERROR: Process_Review_Updates_Middleware(): " + ex.ToString());
                    }
                    Returnvalue = false;
                }
            }

            return Returnvalue;
        }

        /// Set hercust..CustComments status to 1 (Approved)
        /// In future, may add other status update (2: Pending); (3: Not Approved)
        /// Update hercust..CustComments MagentoVetted: -1: Error; 1: Vetted in Magento
        public static Boolean Product_Review_Status_Update(int CustCommentsid = 0)
        {
            Boolean bReturnvalue = true;
            DataTable dt;
            String GetOut;
            String APIGet = "";
            String Status;
            dynamic dJson;
            String APIReturn = "";

            if (CustCommentsid > 0)
            {
                dt = Helper.Sql_Misc_Fetch("SELECT Magentoid, review, vetted, rating, Replace(Replace(Custname, '(male)', ''), '(female)', '') [customername] FROM Hercust..CustComments WHERE id = " + CustCommentsid.ToString());
            }
            else
            {
                dt = Helper.Sql_Misc_Fetch("SELECT Magentoid, review, vetted, rating, Replace(Replace(Custname, '(male)', ''), '(female)', '') [customername] FROM Hercust..CustComments WHERE ISNULL(Magentoid,0) > 0 AND ISNULL(MagentoVetted,0) = 0 AND Vetted = 1 AND convert(varchar(32),ID) NOT IN (SELECT source_id FROM Communications..Middleware WHERE status = 100 AND source_table = 'hercust..custcomments') Order BY id");
            }

            foreach (DataRow dr in dt.Rows)
            {
                APIGet = Helper.MagentoApiGet("V1/reviews/" + dr["magentoid"].ToString());

                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine(APIGet);
                }

                if (APIGet.Length > 4 && APIGet.Substring(0, 5).ToUpper() == "ERROR")
                {
                    Helper.Sql_Misc_NonQuery("UPDATE Hercust..CustComments SET MagentoVetted = -1 WHERE Magentoid = " + dr["magentoid"].ToString());
                }
                else
                {
                    dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(APIGet);

                    if (dJson.review_status != null)
                    {
                        try
                        {
                            Status = dJson.review_status;

                            if (MagetnoProductAPI.DevMode > 0)
                            {
                                Console.WriteLine("Status: " + Status.ToString());
                                Console.WriteLine("Rating: " + dJson.ratings[0].value.ToString());
                            }

                            if (MagetnoProductAPI.DevMode > 0)
                            {
                                Console.WriteLine("Vetted: " + dr["vetted"].ToString());
                            }

                            if (dr["vetted"].ToString() == "1" || dr["vetted"].ToString().ToLower() == "true")
                            {
                                dJson.review_status = 1;        //APPROVED
                            }
                            else
                            {
                                //{
                                //    dJson.review_status = 3;      //NOT APPROVED
                                dJson.review_status = 2;      //PENDING
                            }

                            dJson.title = "Review";
                            dJson.detail = dr["review"].ToString(); //Change for TE's updates

                            //Ratings update 2024-04-23
                            if (dr["rating"].ToString() != dJson.ratings[0].value.ToString())
                            {
                                if (MagetnoProductAPI.DevMode > 0)
                                {
                                    Console.WriteLine("RATING CHG: " + dr["rating"].ToString() + " = " + dJson.ratings[0].value.ToString());
                                }
                                dJson.ratings[0].value = dr["rating"].ToString();
                            }

                            //custname
                            dJson.nickname = dr["customername"].ToString(); 
                           
                            //location NOT IN dJson 
                            //occupation NOT IN dJson 

                            APIGet = Newtonsoft.Json.JsonConvert.SerializeObject(dJson);                           

                            GetOut = @"{""review"":" + APIGet + "}";
                            if (MagetnoProductAPI.DevMode > 0)
                            {
                                Console.WriteLine(APIGet);
                                Console.WriteLine(" -------------------------- ");
                                Console.WriteLine(GetOut);
                                Console.WriteLine(" -------------------------- ");
                            }

                            APIReturn = Helper.MagentoApiPush("PUT", "V1/reviews/" + dr["magentoid"].ToString(), GetOut);

                            if (APIReturn.ToLower().Contains("error"))
                            {
                                if (MagetnoProductAPI.DevMode > 0)
                                {
                                    Console.WriteLine("ERROR: :: " + APIReturn);
                                }
                                Helper.SendEmail("Magento Review Approval FAILED", "Magento Review Approval FAILED: Review magid: " + dr["magentoid"].ToString(), "robot@herroom.com", "thomas@andragroup.com");
                            }
                            else if (dr["vetted"].ToString() == "1" || dr["vetted"].ToString().ToLower() == "true") //approved 
                            {
                                Helper.Sql_Misc_NonQuery("UPDATE Hercust..CustComments SET MagentoVetted = 1 WHERE Magentoid = " + dr["magentoid"].ToString());
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("ERROR: NO Status RETURNED (1): " + ex.ToString());
                        }
                    }
                    else
                    {
                        Console.WriteLine("ERROR: NO Status RETURNED (2): dJson.review_status == null");
                    }
                }
            }

            return bReturnvalue;
        }    

        public static Boolean Product_Reviews(long Middlewareid = 0)
        {
            Boolean bReturnvalue = true;
            DataTable dt = new DataTable();
            DataTable dt2;
            dynamic dJson;
            dynamic dJsonGet;
            int itemCount = 0;
            string Sql = "";
            string GetData;
            dynamic dJsonGetItem;

            try
            {
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, from_magento FROM Communications..Middleware WHERE id = " + Middlewareid.ToString());
                }
                else
                {
                    // INSERT ROW TO DO GET ON
                    Helper.Sql_Misc_NonQuery("EXEC communications..[proc_mag_review_get_row_insert]");

                    dt2 = Helper.Sql_Misc_Fetch("SELECT TOP 1 * FROM Middleware where status = 100 AND source_table = 'get reviews' order by 1 desc");
                    if (dt2.Rows.Count == 1)
                    {
                        //Process GET
                        if (Helper.MagentoApiPush(long.Parse(dt2.Rows[0][0].ToString())).ToUpper() == "OK")
                        {
                            dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, from_magento FROM Communications..Middleware WHERE id = " + dt2.Rows[0][0].ToString());
                        }
                    }
                }

                foreach (DataRow dr in dt.Rows)     // SHOULD BE 1 FOR NOW
                {
              
                    dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(dr["from_magento"].ToString());
 
                    itemCount = dJson.total_count;
                    for (int xx = 0; xx < itemCount; xx++)
                    {
                        if (dJson.items[xx].review_status == 2) //  2 == PENDING
                        {
                             //DO GET FOR EACH REVIEW: https://mcprod.herroom.com/rest/all/V1/reviews/818103
                            GetData = Helper.MagentoApiGet("https://mcprod.herroom.com/rest/all/V1/reviews/" + dJson.items[0].id);

                            if (GetData.Length > 0)
                            {
                                dJsonGet = Newtonsoft.Json.JsonConvert.DeserializeObject(GetData);
                                dJsonGetItem = dJson.items[xx];

                                if (MagetnoProductAPI.DevMode > 0)
                                {
                                    Console.WriteLine(dJsonGetItem.nickname);
                                    Console.WriteLine(dJsonGetItem.customer_id);
                                    Console.WriteLine(dJsonGetItem.created_at);
                                    Console.WriteLine("RATING: " + dJsonGet.ratings[0].value);
                                    Console.WriteLine(dJsonGetItem.detail);
                                    Console.WriteLine(dJsonGetItem.id);
                                    Console.WriteLine(dJsonGetItem.entity_pk_value);
                                }      

                                Sql = "INSERT hercust..custcomments(stylenumber, CustName, email, Location, Occupation, Vetted, RevDate, ItemSize, Rating, AgeGroup, HeightGroup, ForTomima, Augmented, Review, magentoid, magentovetted) "
                                   + "SELECT (SELECT Stylenumber from herroom..styles WHERE magentoid = " + dJson.items[xx].entity_pk_value + ") "
                                   + ", '" + dJsonGetItem.nickname + "', (SELECT isnull(email,'') FROM HERCUST..CUST WHERE magentoid = " + dJsonGetItem.customer_id + ") "
                                   + ", ' ', ' ', 0, '" + dJsonGetItem.created_at + "', 'S' "
                                   + ", REPLACE(" + dJsonGet.ratings[0].value + ",'''','`') "
                                 //+ ",0"  // we do not currently get rating in json:: rating[]
                                   + ", 1, 2, 0, 0 "
                                   + ", '" + dJsonGetItem.detail + "' "
                                   + " , " + dJsonGetItem.id + ", 0 "
                                   + " WHERE " + dJsonGetItem.id + " NOT IN (SELECT ISNULL(magentoid,0) FROM Hercust..CustComments)";

                                Helper.Sql_Misc_NonQuery(Sql);

                                if (MagetnoProductAPI.DevMode > 0)
                                {
                                    Console.WriteLine(" ----------------- ");
                                    Console.WriteLine(Sql);
                                    Console.WriteLine(" ----------------- ");
                                }
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("-- ERROR Product_Reviews(" + Middlewareid.ToString() + "): " + ex.ToString());
                 bReturnvalue = false;
            }

            return bReturnvalue;
        }

        //2024-10-09
        public static Boolean Product_Reviews_GET()
        {
            Boolean bReturnvalue = true;
            DataTable dt = new DataTable();
            //DataTable dt2;
            dynamic dJson;
            dynamic dJsonGet;
            int itemCount = 0;
            string Sql = "";
            string GetData;
            dynamic dJsonGetItem;
            string Datetouse;

            try
            {
                dt = Helper.Sql_Misc_Fetch("SELECT REPLACE(Convert(varchar ,Max(revdate),111),'/','-') + 'T00:00' [revdate] FROM Hercust..CustComments WHERE ISNULL(Magentoid,0) > 0");
                Datetouse = dt.Rows[0]["revdate"].ToString();

                if (dt.Rows.Count == 1)
                {
                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("V1/reviews/?searchCriteria[filter_groups][0][filters][0][field]=created_at&searchCriteria[filter_groups][0][filters][0][value]=" + Datetouse + "&searchCriteria[filter_groups][0][filters][0][condition_type]=gt");
                    }

                    GetData = Helper.MagentoApiGet("V1/reviews/?searchCriteria[filter_groups][0][filters][0][field]=created_at&searchCriteria[filter_groups][0][filters][0][value]=" + Datetouse + "&searchCriteria[filter_groups][0][filters][0][condition_type]=gt");

                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("DT: " + Datetouse);
                        Console.WriteLine(GetData);
                    }

                    dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(GetData);

                    itemCount = dJson.total_count;
                    for (int xx = 0; xx < itemCount; xx++)
                    {
                        if (dJson.items[xx].review_status == 2) //  2 == PENDING
                        {
                            //DO GET FOR EACH REVIEW: https://mcprod.herroom.com/rest/all/V1/reviews/818103
                            GetData = Helper.MagentoApiGet("https://mcprod.herroom.com/rest/all/V1/reviews/" + dJson.items[0].id);

                            if (GetData.Length > 0)
                            {
                                dJsonGet = Newtonsoft.Json.JsonConvert.DeserializeObject(GetData);
                                dJsonGetItem = dJson.items[xx];

                                if (MagetnoProductAPI.DevMode > 0)
                                {
                                    Console.WriteLine(dJsonGetItem.nickname);
                                    Console.WriteLine(dJsonGetItem.customer_id);
                                    Console.WriteLine(dJsonGetItem.created_at);
                                    Console.WriteLine("RATING: " + dJsonGet.ratings[0].value);
                                    Console.WriteLine(dJsonGetItem.detail);
                                    Console.WriteLine(dJsonGetItem.id);
                                    Console.WriteLine(dJsonGetItem.entity_pk_value);
                                }

                                Sql = "INSERT hercust..custcomments(stylenumber, CustName, email, Location, Occupation, Vetted, RevDate, ItemSize, Rating, AgeGroup, HeightGroup, ForTomima, Augmented, Review, magentoid, magentovetted) "
                                   + "SELECT (SELECT Stylenumber from herroom..styles WHERE magentoid = " + dJson.items[xx].entity_pk_value + ") "
                                   + ", '" + dJsonGetItem.nickname + "', (SELECT isnull(email,'') FROM HERCUST..CUST WHERE magentoid = " + dJsonGetItem.customer_id + ") "
                                   + ", ' ', ' ', 0, '" + dJsonGetItem.created_at + "', 'S' "
                                   + ", REPLACE(" + dJsonGet.ratings[0].value + ",'''','`') "
                                   //+ ",0"  // we do not currently get rating in json:: rating[]
                                   + ", 1, 2, 0, 0 "
                                   + ", '" + dJsonGetItem.detail + "' "
                                   + " , " + dJsonGetItem.id + ", 0 "
                                   + " WHERE " + dJsonGetItem.id + " NOT IN (SELECT ISNULL(magentoid,0) FROM Hercust..CustComments)";

                                Helper.Sql_Misc_NonQuery(Sql);

                                if (MagetnoProductAPI.DevMode > 0)
                                {
                                    Console.WriteLine(" ----------------- ");
                                    Console.WriteLine(Sql);
                                    Console.WriteLine(" ----------------- ");
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
                    Console.WriteLine("-- ERROR Product_Reviews_Get(): " + ex.ToString());
                }
                bReturnvalue = false;
            }

            return bReturnvalue;
        }

        public static Boolean Products_Reviews_2(int DaysinthePast = -2, int RowProcessingLimit = 2)
        {
            string GetData;
            dynamic dJson;
            int itemCount = 0;
            DateTime Now;  
            string WhentoUse;
            bool bReturnValue = true;
            int Processed = 0;

            try
            {
                if (DaysinthePast > 0)
                {
                    Now = DateTime.Now.AddDays(-DaysinthePast);
                }
                else
                {
                    Now = DateTime.Now.AddDays(DaysinthePast);
                }

                WhentoUse = Now.ToString("yyyy-MM-dd") + "T00:00";
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine(WhentoUse);
                }

                GetData = Helper.MagentoApiGet("https://mcprod.herroom.com/rest/all/V1/reviews/?searchCriteria[filter_groups][0][filters][0][field]=created_at&searchCriteria[filter_groups][0][filters][0][value]=" + WhentoUse + "Z&searchCriteria[filter_groups][0][filters][0][condition_type]=gt");

                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine(GetData);
                }

                dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(GetData);

                itemCount = dJson.total_count;

                for (int xx = 0; xx < itemCount; xx++)
                {
                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine(dJson.items[xx].id + "; " + dJson.items[xx].entity_pk_value);
                    }

                    if (Processed < RowProcessingLimit)
                    {
                        if (HTCustomer_Comments_Insert(dJson.items[xx].id.ToString()))
                        {
                            Processed++;
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine("Products_Reviews_2: " + ex.ToString());
                }
                bReturnValue = false;
            }
                return bReturnValue;
        }

        //WORKS, may be used in above Block()   
        private static Boolean HTCustomer_Comments_Insert(string ReviewId) //(dynamic dJsonReview, dynamic dJsonItem)
        {
            string Sql = "";
            string GetData;
            dynamic dJsonGet;
            DataTable dt;
            string ReviewDetail;
            string custName;        // 2024-04-24

            dt = Helper.Sql_Misc_Fetch("SELECT COUNT(*) [count] FROM Hercust..CustComments WHERE magentoid = " + ReviewId);
            if (dt.Rows[0]["count"].ToString() == "0")
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine("MISSING REVIEW :: " + ReviewId);
                }

                GetData = Helper.MagentoApiGet("https://mcprod.herroom.com/rest/all/V1/reviews/" + ReviewId);

                if (GetData.Length > 0)
                {
                    dJsonGet = Newtonsoft.Json.JsonConvert.DeserializeObject(GetData);

                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine(dJsonGet.nickname);
                        Console.WriteLine(dJsonGet.customer_id);
                        Console.WriteLine(dJsonGet.created_at);
                        Console.WriteLine("RATING: " + dJsonGet.ratings[0].value);
                        Console.WriteLine(dJsonGet.detail);
                        Console.WriteLine(dJsonGet.id);                     // review id 
                        Console.WriteLine(dJsonGet.entity_pk_value);        // herroom..styles.magentoid
                    }

                    ReviewDetail = dJsonGet.detail.ToString().Replace("'", "''");
                    ReviewDetail = ReviewDetail.Length > 300 ? ReviewDetail.Substring(0, 299) : ReviewDetail;
                    //No Gender in GET detail
                    //custName = dJsonGet.store_id == "2" ? custName + "(Herroom dJsonGet.nickname + " (" + d

                    Sql = "INSERT hercust..custcomments(stylenumber, CustName, email, Location, Occupation, Vetted, RevDate, ItemSize, Rating, AgeGroup, HeightGroup, ForTomima, Augmented, Review, magentoid, magentovetted) "
                       + "SELECT (SELECT TOP 1 Stylenumber from herroom..styles WHERE magentoid = " + dJsonGet.entity_pk_value + ") "
                       + ", '" + dJsonGet.nickname + "', (SELECT TOP 1 isnull(email,'') FROM HERCUST..CUST WHERE magentoid = " + dJsonGet.customer_id + ") "
                       + ", ' ', ' ', 0, '" + dJsonGet.created_at + "', 'S' "
                       + ", REPLACE(" + dJsonGet.ratings[0].value + ",'''','`') "
                       + ", 1, 2, 0, 0 "
                       + ", '" + ReviewDetail + "' "
                       + " , " + dJsonGet.id + ", 0 "
                       + " WHERE " + dJsonGet.id + " NOT IN (SELECT ISNULL(magentoid,0) FROM Hercust..CustComments)";

                    Helper.Sql_Misc_NonQuery(Sql);

                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine(" ----------------- ");
                        Console.WriteLine(Sql);
                        Console.WriteLine(" ----------------- ");
                    }
                }
            }
            else
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// 2024-02-07, new review form info coming into middleware_incoming to use that
        /// Still need to get Product.magentoid from API GET on specific review (for now)
        /// <param name="Middlewareid"></param>
        /// <param name="MaxrowstoProcess"></param>
        /// <returns></returns>
        public static Boolean Process_Reviews_MiddlewareIncoming(long Middlewareid = 0, int MaxrowstoProcess = 2)
        {
            string Sql;
            string GetData;
            dynamic dJson;
            dynamic dJsonGet;
            DataTable dt;
            string ReviewDetail;
            string Reviewid;
            string HeightGroup;
            string AgeGroup;
            string CustName;
            int iAge = 1;

            try
            {
                Sql = "SELECT TOP " + MaxrowstoProcess.ToString() + " id, posted, from_magento, endpoint_method FROM communications..middleware_incoming WHERE endpoint_method = 'incoming - module form' AND from_magento like '%review form%' ";

                if (Middlewareid > 0)
                {
                    Sql += " AND id = " + Middlewareid.ToString();
                }
                else
                {
                    Sql += " AND status = 0 AND posted <= getdate() ";  //will probably never post in the future really
                }

                Sql += " ORDER BY ID ";

                dt = Helper.Sql_Misc_Fetch(Sql);
                Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                {
                    try
                    {
                        dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(dr["from_magento"].ToString());
                        Reviewid = dJson.form.entity_id;

                        GetData = Helper.MagentoApiGet("https://mcprod.herroom.com/rest/all/V1/reviews/" + Reviewid);
                        dJsonGet = Newtonsoft.Json.JsonConvert.DeserializeObject(GetData);

                        //entity_pk_value is the only way to get style's magentoid
                        if (MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine(Reviewid + "; " + dJsonGet.entity_pk_value);
                        }

                        ReviewDetail = dJson.form.comment.ToString().Replace("'", "''");
                        ReviewDetail = ReviewDetail.Length > 300 ? ReviewDetail.Substring(0, 299) : ReviewDetail;

                        HeightGroup = "2";
                        if (dJson.form.height.ToString().ToLower().Contains("tall"))
                        {
                            HeightGroup = "1";
                        }
                        else if (dJson.form.height.ToString().ToLower().Contains("petite"))
                        {
                            HeightGroup = "3";
                        }

                        AgeGroup = "1";     //Age:  teen, 20s, 30s, 40s, 50s, 60s, 70s, 80 & over
                                            //sAge = dJson.form.age;
                        if (dJson.form.age.ToString().Length > 0 && int.TryParse(dJson.form.age.ToString().Substring(0, 1), out iAge))
                        {
                            AgeGroup = iAge.ToString();
                        }

                        CustName = dJson.form.customer_name.ToString().Length > 70 ? CustName = dJson.form.customer_name.ToString().Substring(0, 70) : dJson.form.customer_name.ToString();
                        CustName += " (" + dJson.form.gender + ")";

                        Sql = "EXEC Communications..[proc_mag_review_middleware_insert] @StyleMagentoid = " + dJsonGet.entity_pk_value
	                        + ", @CustName = '" + CustName + "' "
                            + ", @Email = '" + dJson.form.customer_email + "' "
                            + ", @Location = '" + dJson.form.address + "' "
                            + ", @Occupation = '" + dJson.form.occupation + "' "
                            + ", @RevDate = '" + dJsonGet.created_at + "' "
                            + ", @Size = '" + dJson.form.size + "' "
                            + ", @RatingValue = " + dJson.form.rating_value
                            + ", @Agegroup = " + AgeGroup
                            + ", @HeightGroup = " + HeightGroup
                            + ", @Review = '" + ReviewDetail + "' "
                            + ", @ReviewMagentoid = " + dJsonGet.id;

                        if (MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine(" ----------------- ");
                            Console.WriteLine(Sql);
                            Console.WriteLine(" ----------------- ");
                        }

                        if (MagetnoProductAPI.DevMode < 2)
                        {
                            Helper.Sql_Misc_NonQuery(Sql);
                            Helper.Sql_Misc_NonQuery("UPDATE communications..middleware_incoming SET status = 600 WHERE id = " + dr["id"].ToString());
                        }
                    }
                    catch (Exception exInner)
                    {
                        if (MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine("ERROR Process_Reviews_MiddlewareIncoming(" + Middlewareid.ToString() + ":: " + exInner.ToString());
                        }
                        Helper.Sql_Misc_NonQuery("UPDATE communications..middleware_incoming SET status = 700, Error_message='ERROR' WHERE id = " + dr["id"].ToString());
                    }
                }

            }
            catch(Exception ex)
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine("ERROR Process_Reviews_MiddlewareIncoming(" + Middlewareid.ToString() + ":: " + ex.ToString());
                }
                return false;
            }

            return true;
        }

        public static Boolean Process_Get_Color_Options()
        {
            DataTable dt;
            DataTable dtHC;

            //CREATE THIS IN A 1/day SQL JOB
            //Helper.Sql_Misc_NonQuery("IF 0 = (SELECT COUNT(*) FROM Communications..Middleware WHERE endpoint_method = 'get' AND source_table = 'herroom..colors' AND Status = 100) " 
            //    + " INSERT Communications..Middleware(posted, status, worker_id, endpoint_name, endpoint_method, source_table, source_id) "
            //    + " SELECT getdate(), 100, 100, 'V1/products/attributes/color/options', 'GET', 'herroom..colors', 'all' ;");

            dt = Helper.Sql_Misc_Fetch("SELECT TOP 1 id FROM Communications..Middleware WHERE endpoint_method = 'get' AND source_table = 'herroom..colors' AND Status = 100 ORDER BY id DESC");

            if (dt.Rows.Count == 1)
            {
                Helper.MagentoApiPush(long.Parse(dt.Rows[0]["id"].ToString()));
            }

            dtHC = Helper.Sql_Misc_Fetch("SELECT TOP 1 id FROM Communications..Middleware WHERE endpoint_method = 'get' AND source_table = 'her_color' AND Status = 100 ORDER BY id DESC");

            if (dtHC.Rows.Count == 1)
            {
                Helper.MagentoApiPush(long.Parse(dtHC.Rows[0]["id"].ToString()));
            }

            return true;
        }

        public static Boolean RerunSimple700()
        {
            DataTable dt;

            dt = Helper.Sql_Misc_Fetch("SELECT top 500 id FROM middleware WHERE source_table = 'herroom..items' AND posted > '2023-09-28' AND status = 700");
            foreach (DataRow dr in dt.Rows)
            {
                Process_Product_Simple_Middleware(long.Parse(dr["id"].ToString()));
            }
            return true;
        }

        public static Boolean Process_Products_Active_NotinMagento()
        {
            DataTable dtStyles;

            //Herroom..Styles
            dtStyles = Helper.Sql_Misc_Fetch("SELECT StyleNumber FROM herroom..styles SS with (nolock) "
                + " WHERE SS.Active = 1 AND ss.Magentoid IS NULL "
                + " AND 0 < (SELECT COUNT(*) FROM Herroom..items II WHERE II.StyleNumber = SS.StyleNumber AND II.Active = 1) "
                + " AND 0 = (SELECT COUNT(*) FROM Communications..Middleware WHERE source_table = 'herroom..styles' AND source_id = StyleNumber AND status = 100) "
                + " ORDER BY 1;");

            foreach (DataRow drStyles in dtStyles.Rows)
            {
                Helper.Sql_Misc_NonQuery("EXEC Communications..[proc_mag_product_configurable_json_post_api] @Stylenumber = '" + drStyles["stylenumber"].ToString() + "', @BlankProductLinks = 1;");
            }

            //Herroom..Items
            Helper.Sql_Misc_NonQuery("EXEC Communications..[proc_mag_get_product_simple_active_notinmagento]");

            return true;
        }

        /// <summary>
        /// UpdateType must = herroom..productlinks OR herroom..stylessizingchart OR herroom..stylecategories
        /// 
        /// </summary>
        /// <param name="Middlewareid"></param>
        /// <param name="StyleNumber"></param>
        /// <param name="UpdateType"></param>
        /// <returns></returns>
        public static Boolean Process_Configurable_Updates(long Middlewareid = 0, string StyleNumber = "", string UpdateType = "", string Status = "100", int MaxRowstoProcess=0)
        {
            Boolean bReturnvalue = true;
            String SqlString;
            DataTable dt = new DataTable();
            String APIReturn = "";

            try
            {
                if (MaxRowstoProcess > 0)
                {
                    SqlString = "SELECT TOP " + MaxRowstoProcess.ToString() + " id, source_id, endpoint_method, source_table FROM Communications..Middleware with (nolock) WHERE source_table IN ('herroom..productlinks','herroom..stylessizingchart','herroom..stylecategories', 'herroom..stylevideo', 'herroom..stylevideos', 'herroom..styleproductlinks', 'herroom..stylefeatures', 'herroom..stylekeywords', 'herroom..stylelabels', 'herroom..stylenewarrival', 'herroom..itemsexpecteddate', 'herroom..stylematching') AND to_magento IS NOT NULL ";
                }
                else
                {
                    SqlString = "SELECT id, source_id, endpoint_method, source_table FROM Communications..Middleware with (nolock) WHERE source_table IN ('herroom..productlinks','herroom..stylessizingchart','herroom..stylecategories', 'herroom..stylevideo', 'herroom..stylevideos', 'herroom..styleproductlinks', 'herroom..stylefeatures', 'herroom..stylekeywords', 'herroom..stylelabels', 'herroom..stylenewarrival', 'herroom..itemsexpecteddate', 'herroom..stylematching') AND to_magento IS NOT NULL ";
                }

                if (Middlewareid > 0) 
                {
                    SqlString += " AND id = " + Middlewareid.ToString();
                }
                else if (StyleNumber != "")
                {
                    SqlString += " AND source_id = '" + StyleNumber + "'";
                }
                else if (UpdateType != "")
                {
                    SqlString += " AND source_table = '" + UpdateType + "'";
                }

                if (Status != "")
                {
                    SqlString += " AND Posted <= Getdate() AND Status = " + Status ;
                }
                else
                {
                    SqlString += " AND Posted <= Getdate() AND Status = 100 AND to_magento IS NOT NULL";
                }

                SqlString += " ORDER BY Posted, id";

                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine(SqlString);
                }

                dt = Helper.Sql_Misc_Fetch(SqlString);

                Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                {
                    //MStaging Load ONLY 2023-12-28
                    APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");        
 
                    if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                    {
                        //Middleware should already by set to 500
                        Console.WriteLine("ERROR for Configurable Update: " + dr["source_id"].ToString() + "; source_table: " + dr["source_table"].ToString());
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, from_magento = '" + APIReturn + "' WHERE ID = " + dr["id"].ToString());

                        Helper.SendEmail("ERROR for Configurable Update: " + dr["source_id"].ToString(), "ERROR for Configurable Update: " + dr["source_id"].ToString() + "; ID: " + dr["id"].ToString() + "; source_table: " + dr["source_table"].ToString(), "thomas@andragroup.com", "alerts@herroom.com");
                    }
                    else //Should return style magentoid that we do not need
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = 'OKAY' WHERE ID = " + dr["id"].ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Process_Configurable_Updates(" + Middlewareid.ToString() + ": " + ex.ToString());
                bReturnvalue = false;
            }

            return bReturnvalue;
        }


        public static Boolean ReviewTEST()
        {
            string GetData;
            dynamic dJsonGet;
            //dynamic dJsonGetItem;

            GetData = Helper.MagentoApiGet("https://mcprod.herroom.com/rest/all/V1/reviews/926077");

            if (GetData.Length > 0)
            {
                Console.WriteLine(GetData);
                dJsonGet = Newtonsoft.Json.JsonConvert.DeserializeObject(GetData);
                Console.WriteLine("RATING: " + dJsonGet.ratings[0].value);
            }

            return true;
        }

        public static Boolean Product_Children(long Middlewareid = 0, int MaxRowstoProcess = 10)
        {
            dynamic dJson;
            DataTable dt;
            DataTable dtStyle;
            DataTable dtSkus;
            DataTable dtList;
            String MailOutput = @"<table border=""1""><tr><th>Stylenumber</th><th>Magneto Lins</th><th>HT Sku Count</th><th>Relink Staged</th></tr>";
            int IssuesFound = 0;

            Hashtable HT = new System.Collections.Hashtable();

            try
            {
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT status, id, source_id FROM communications..middleware WHERE id = " + Middlewareid.ToString() + " AND source_table = 'herroom..stylelinks' ");
                }
                else
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT TOP " + MaxRowstoProcess.ToString() + " status, id, source_id FROM communications..middleware WHERE Posted <= GETDATE() AND source_table = 'herroom..stylelinks' AND endpoint_method = 'GET' AND Status = 100 ORDER BY id");
                }

                Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                {
                    if (dr["status"].ToString() == "100")
                    {
                        Helper.MagentoApiPush(long.Parse(dr["id"].ToString()));
                    }

                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("SELECT id, ISNULL(from_magento, '') [from_magento] FROM Communications..Middleware WHERE id = " + dr["id"].ToString());
                    }
                    dtList = Helper.Sql_Misc_Fetch("SELECT id, ISNULL(from_magento, '') [from_magento] FROM Communications..Middleware WHERE id = " + dr["id"].ToString());

                    Helper.Sql_Misc_NonQuery("UPDATE Communications..middleware SET status=600 WHERE id = " + dr["id"].ToString());

                    dtStyle = Helper.Sql_Misc_Fetch("SELECT COUNT(*) [skus] FROM Herroom..Items WHERE stylenumber = '" + dr["source_id"].ToString() + "' AND Magentoid IS NOT NULL AND Active = 1;");

                    dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(dtList.Rows[0]["from_magento"].ToString());
                    //dJson.Count            sku and id work 

                    if (dJson.Count == null)
                    {
                        continue;
                    }

                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine(dJson.Count.ToString() + "; Style skus: " + dtStyle.Rows[0]["skus"].ToString());
                    }

                    if (dJson.Count < int.Parse(dtStyle.Rows[0]["skus"].ToString()))
                    {
                        IssuesFound++;

                        if (dJson.Count > 0)
                        {
                            //LOAD HASHTable for faster searching
                            HT.Clear();
                            for (int xx = 0; xx < dJson.Count; xx++)
                            {
                                HT.Add(xx, dJson[xx].id.ToString());
                            }

                            dtSkus = Helper.Sql_Misc_Fetch("SELECT magentoid FROM Herroom..Items WHERE Stylenumber = '" + dr["source_id"].ToString() + "' AND active = 1 AND magentoid IS NOT NULL");
                            foreach (DataRow drSkus in dtSkus.Rows)
                            {
                                //if srSku magentoid NOT IN dJson[xx].id, insert [proc_mag_product_links_json_api_child_sku]
                                if (!HT.ContainsValue(drSkus["magentoid"].ToString()))
                                {
                                    if (MagetnoProductAPI.DevMode > 0)
                                    {
                                        Console.WriteLine("MISSING LINK: " + drSkus["magentoid"].ToString());
                                    }
                                    Helper.Sql_Misc_NonQuery("EXEC Communications..[proc_mag_product_links_json_api_child_sku] @Itemid=0, @Magentoid=" + drSkus["magentoid"].ToString());
                                }
                            }
   
                            MailOutput += "<tr><td>" + dr["source_id"].ToString() + "</td><td>" + dJson.Count.ToString() + "</td><td>" + dtStyle.Rows[0]["skus"].ToString() + "</td><td>YES</td></tr>";
                        }
                        else
                        {
                            Helper.Sql_Misc_NonQuery("EXEC Communications..[proc_mag_product_options_json_api] '" + dr["source_id"].ToString() + "' ;");
                            Helper.Sql_Misc_NonQuery("EXEC Communications..[proc_mag_product_links_json_api_child] @Stylenumber = '" + dr["source_id"].ToString() + "', @RuninMinutes=15, @reCreate=1 ;");

                            MailOutput += "<tr><td>" + dr["source_id"].ToString() + "</td><td>" + dJson.Count.ToString() + "</td><td>" + dtStyle.Rows[0]["skus"].ToString() + "</td><td>ALL</td></tr>";
                        }
                    }
                }

                MailOutput += "</table>";

                if (IssuesFound > 0)
                {
                    Helper.SendEmail("Magento Style-Sku Links Missing = " + IssuesFound.ToString(), MailOutput, "thomas@andragroup.com");
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("ERROR: " + ex.ToString());
                return false;
            }

            return true;
        }


        //2024-05-29- NEW for SIMPLE URLKEY UPDATES FROM 
        public static Boolean Process_Product_Simple_ItemsUrlKey(long Middlewareid = 0, int MaxRowstoProcess = 10, string Environment = "")
        {
            Boolean bReturnvalue = true;
            DataTable dt = new DataTable();
            String APIReturn = "";
            long mid;

            try
            {
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, ISNULL(to_magento,'') [to_magento], source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE source_table = 'Herroom..ItemsUrlKey' AND id = " + Middlewareid.ToString() + " AND Status = 100");
                }
                else
                {
                    if (MaxRowstoProcess > 0)
                    {
                        dt = Helper.Sql_Misc_Fetch("SELECT TOP " + MaxRowstoProcess.ToString() + " id, ISNULL(to_magento,'') [to_magento], source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND source_table = 'Herroom..ItemsUrlKey' ORDER BY ID");
                    }
                    else
                    {
                        dt = Helper.Sql_Misc_Fetch("SELECT id, ISNULL(to_magento,'') [to_magento], source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE Posted <= Getdate() AND Status=100 AND source_table = 'Herroom..ItemsUrlKey' ORDER BY ID");
                    }
                }

                //Helper.Middleware_Status_Update(dt, 200);

                foreach (DataRow dr in dt.Rows)
                {
                    Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET Status = 200 WHERE ID = " + dr["id"].ToString()); ;

                    if (Environment == "")
                    {
                        APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString())).Replace("\"", "");
                    }
                    else
                    {
                        //APIReturn = Helper.MagentoApiPush(long.Parse(dr["id"].ToString()), Environment);
                        APIReturn = Helper.MagentoApiPushStaging(long.Parse(dr["id"].ToString()));
                    }

                    APIReturn = APIReturn.Replace("'", "`");

                    if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                    {
                        //Middleware should already by set to 500
                        Console.WriteLine("ERROR for Process_Product_Simple_ItemsUrlKey: " + dr["source_id"].ToString());

                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, from_magento = '" + APIReturn + "', Module = '" + Environment + "' WHERE ID = " + dr["id"].ToString());

                        //Helper.SendEmail("ERROR for Process_Product_Simple_ItemsUrlKey: " + dr["source_id"].ToString(), "ERROR for Color Processing: " + dr["source_id"].ToString() + "; ID: " + dr["id"].ToString(), "thomas@andragroup.com", "alerts@herroom.com");
                    }
                    else if(APIReturn.ToLower().Contains("the consumer isn"))
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=700, from_magento = '" + APIReturn + "', Module = '" + Environment + "' WHERE ID = " + dr["id"].ToString());

                    }
                    else //Should return new value
                    {
                        if (long.TryParse(APIReturn, out mid) && dr["endpoint_method"].ToString().ToUpper() == "POST")
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + mid.ToString() + "', Module = '" + Environment + "' WHERE ID = " + dr["id"].ToString());

                        }
                        else
                        {
                            Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET status=600, from_magento = '" + APIReturn + "', Module = '" + Environment + "' WHERE ID = " + dr["id"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Process_Product_Simple_ItemsUrlKey(" + Middlewareid.ToString() + ": " + ex.ToString());
                bReturnvalue = false;
            }

            return bReturnvalue;
        }
    }

    class Audit
    {
        public static Boolean Her_Color_Audit(long Middlwareid)
        {
            Boolean bReturnvalue = true;
            DataTable dt;
            DataTable dtColors;
            DataTable dtColorsHT;
            DataRow dr;
            DataRow[] drColors;
            dynamic dJson;
            dynamic dJsonColor;
            Boolean drColorHTFound;

            try
            {
                dtColors = Helper.Sql_Misc_Fetch("SELECT colorcode, ISNULL(Magentoid,0) [magentoid], ISNULL(Hercolorid,0) [hercolorid] FROM Herroom..Colors ");
                dtColors.PrimaryKey = new DataColumn[] { dtColors.Columns["colorcode"] };

                dt = Helper.Sql_Misc_Fetch("SELECT id, from_magento FROM Communications..Middleware WHERE source_table = 'her_color' AND id=" + Middlwareid.ToString());

                if (dt.Rows.Count > 0)
                {
                    dr = dt.Rows[0];

                    dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(dr["from_magento"].ToString());

                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("COUNT: " + dJson.Count.ToString());
                    }

                    //FIRST SEARCH FOR MISSING HERTOOLS ENTRIES 
                    for (int xx = 0; xx < dJson.Count; xx++)
                    {
                        dJsonColor = dJson[xx];

                        if (dJsonColor.label != null && dJsonColor.value != "")
                        {
                            drColors = dtColors.Select("colorcode='" + dJsonColor.label + "'");

                            if (drColors != null && drColors.Length > 0)
                            {
                                if (MagetnoProductAPI.DevMode > 1)
                                {
                                    Console.WriteLine(dJsonColor.label + "; " + dJsonColor.value + " = " + drColors[0]["colorcode"].ToString() + "; " + drColors[0]["hercolorid"].ToString());
                                    Console.WriteLine("----");
                                }

                                if (dJsonColor.value != drColors[0]["hercolorid"].ToString())
                                {
                                    Console.WriteLine("ERROR");
                                    Console.WriteLine(dJsonColor.label + "; " + dJsonColor.value + " = " + drColors[0]["colorcode"].ToString() + "; " + drColors[0]["hercolorid"].ToString());
                                    Console.WriteLine("----");
                                }
                            }
                        }
                    }

                    //SECOND, SERACH FOR MISSING/ERRANT Magento ENTRIES
                    dtColorsHT = Helper.Sql_Misc_Fetch("SELECT colorcode, ISNULL(Magentoid,0) [magentoid], ISNULL(Hercolorid,0) [hercolorid] FROM Herroom..Colors WHERE ISNULL(Hercolorid,0) > 0");
                    //1st try is a bubblesort, fix later
                    foreach (DataRow drColorHT in dtColorsHT.Rows)
                    {
                        drColorHTFound = false;
                        for (int xx = 0; xx < dJson.Count; xx++)
                        {
                            if (!drColorHTFound)
                            {
                                dJsonColor = dJson[xx];

                                if (dJsonColor.value == drColorHT["hercolorid"].ToString())
                                {
                                    drColorHTFound = true;
                                }
                            }
                        }
                        if (!drColorHTFound)
                        {
                            Console.WriteLine("MISSING: " + drColorHT["colorcode"].ToString() + "; ID: " + drColorHT["hercolorid"].ToString());
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine(ex.ToString());
                }
                bReturnvalue = false;
            }

            return bReturnvalue;
        }
    

       public static Boolean Her_Size_Audit(long Middlwareid)
        { 
            Boolean bReturnvalue = true;
            DataTable dt;
            DataTable dtSizes;
            DataTable dtSizesHT;
            DataRow dr;
            DataRow[] drSizes;
            dynamic dJson;
            dynamic dJsonSize;
            Boolean drSizeHTFound;

            try
            {
                dtSizes = Helper.Sql_Misc_Fetch("SELECT size, ISNULL(Magentoid,0) [magentoid], ISNULL(Hersizeid,0) [hersizeid] FROM Herroom..Mastersize ");
                dtSizes.PrimaryKey = new DataColumn[] { dtSizes.Columns["size"] };

                dt = Helper.Sql_Misc_Fetch("SELECT id, from_magento FROM Communications..Middleware WHERE id=" + Middlwareid.ToString());

                if (dt.Rows.Count > 0)
                {
                    dr = dt.Rows[0];

                    dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(dr["from_magento"].ToString());

                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("COUNT: " + dJson.Count.ToString());
                    }
                    
                    //FIRST SEARCH FOR MISSING HERTOOLS ENTRIES 
                    for (int xx = 0; xx < dJson.Count; xx++)
                    {
                        dJsonSize = dJson[xx];

                        if (dJsonSize.label != null && dJsonSize.value != "")
                        {
                            drSizes = dtSizes.Select("size='" + dJsonSize.label + "'");
                            
                            if (drSizes != null && drSizes.Length > 0)
                            {
                                if (MagetnoProductAPI.DevMode > 1)
                                {
                                    Console.WriteLine(dJsonSize.label + "; " + dJsonSize.value + " = " + drSizes[0]["size"].ToString() + "; " + drSizes[0]["hersizeid"].ToString());
                                    Console.WriteLine("----");
                                }

                                if (dJsonSize.value != drSizes[0]["hersizeid"].ToString())
                                {
                                    Console.WriteLine(dJsonSize.label + "; " + dJsonSize.value + " = " + drSizes[0]["size"].ToString() + "; " + drSizes[0]["hersizeid"].ToString());
                                    Console.WriteLine("----");
                                }
                            }
                        }
                    }

                    //SECOND, SERACH FOR MISSING/ERRANT Magento ENTRIES
                    dtSizesHT = Helper.Sql_Misc_Fetch("SELECT size, ISNULL(Magentoid,0) [magentoid], ISNULL(Hersizeid,0) [hersizeid] FROM Herroom..Mastersize WHERE ISNULL(Hersizeid,0) > 0");
                    //1st try is a bubblesort, fix later
                    foreach (DataRow drSizeHT in dtSizesHT.Rows)
                    {
                        drSizeHTFound = false;
                        for (int xx = 0; xx < dJson.Count; xx++)
                        {
                            if (!drSizeHTFound)
                            {
                                dJsonSize = dJson[xx];

                                if (dJsonSize.value == drSizeHT["hersizeid"].ToString())
                                {
                                    drSizeHTFound = true;
                                }
                            }
                        }
                        if (!drSizeHTFound)
                        {
                            Console.WriteLine("MISSING: " + drSizeHT["size"].ToString() + "; ID: " + drSizeHT["hersizeid"].ToString());
                        }
                    }
                }

            }
            catch(Exception ex)
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine(ex.ToString());
                }
                bReturnvalue = false;
            }

            return bReturnvalue;
        }

    }

    // endpoint_method: GET, endpoint_name: V1/stockItems/lowStock?scopeId=0&qty=100&pageSize=1000 
    // Need 3 reports: 1) qty/Bo off 2) magentoid not in HT 3) HT magid not in Magento 
    // RUN THESE IS SEPERATE FUNCTION FOR NOW
    class Inventory
    {
        public static Boolean Process_Low_Inventory_Report(long Middlewareid)
        {
            Boolean bReturnvalue = true;
            DataTable dt;
            DataTable dtHertools;
            string FromMagento;
            dynamic dJson;
            dynamic dJsonItem;
            DataRow[] drHerTools;
            string magisinstock;

            try
            {
                dt = Helper.Sql_Misc_Fetch("SELECT id, from_magento FROM communications..middleware WHERE id = " + Middlewareid.ToString() + " AND source_table = 'stockItems/lowStock' AND endpoint_method = 'GET' AND Status = 600");

                if (dt.Rows.Count > 0)
                {
                    FromMagento = dt.Rows[0]["from_magento"].ToString();

                    //Load table with entire Hertools Items list (active only for now)
                    dtHertools = Helper.Sql_Misc_Fetch("SELECT II.magentoid, qtyonhand, CASE WHEN SS.CloseOut=1 OR II.UpcCloseOut=1 THEN 0 ELSE 1 END [htbo], CASE WHEN (SS.CloseOut=1 OR II.UpcCloseOut=1) AND II.QtyOnHand=0 THEN 0 WHEN SS.Active=1 AND II.Active=1 THEN 1 ELSE 0 END [isinstock] FROM herroom..items II with (nolock) INNER JOIN Herroom..styles SS with (nolock) ON SS.stylenumber = II.StyleNumber AND SS.Active = 1 AND ISNULL(SS.Magentoid, 0) > 0 WHERE II.Active = 1 AND ISNULL(II.Magentoid, 0) > 0");

                    dtHertools.PrimaryKey = new DataColumn[] { dtHertools.Columns["magentoid"] };

                    dJson = Newtonsoft.Json.JsonConvert.DeserializeObject(FromMagento);

                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("COUNT: " + dJson.items.Count.ToString());
                    }

                    for (int xx = 0; xx < dJson.items.Count; xx++)  
                    {
                        dJsonItem = dJson.items[xx];    

                        if (xx % 1000 == 0 && MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine(xx.ToString());
                        }

                        //Find hertools entry for Magento sku; NOTE: Style skus (configurables) will be skipped
                        drHerTools = dtHertools.Select("magentoid=" + dJsonItem.product_id);
                        if (drHerTools.Length > 0)
                        {
                            if (MagetnoProductAPI.DevMode > 1)
                            {
                                Console.WriteLine(dJsonItem.ToString() + "; " + drHerTools[0]["isinstock"].ToString());
                                Console.WriteLine("----");
                            }

                            //"is_in_stock":true,
                            magisinstock = "1";
                            if (dJsonItem.is_in_stock == "false")
                            {
                                magisinstock = "0";        
                            }

                            //if (drHerTools[0]["qtyonhand"].ToString() != dJsonItem.qty.ToString() || drHerTools[0]["htbo"].ToString() != dJsonItem.backorders.ToString())
                            if (drHerTools[0]["qtyonhand"].ToString() != dJsonItem.qty.ToString() || drHerTools[0]["htbo"].ToString() != dJsonItem.backorders.ToString() || drHerTools[0]["isinstock"].ToString() != magisinstock)
                            {
                                //Write line to DB, the updates will be processed later in another function()
                                Helper.Sql_Misc_NonQuery("EXEC Communications..[proc_mag_inventory_issues_insert] @magentoid = " + drHerTools[0]["magentoid"].ToString()
                                + " , @magentoqty = " + dJsonItem.qty.ToString()
                                + " , @qtyonhand = " + drHerTools[0]["qtyonhand"].ToString()
                                + " , @backorders = " + dJsonItem.backorders.ToString()
                                + " , @htbo = " + drHerTools[0]["htbo"].ToString()
                                + " , @magisinstock = " + magisinstock
                                + " , @htactive = " + drHerTools[0]["isinstock"].ToString());

                                if (MagetnoProductAPI.DevMode > 1)
                                {
                                    Console.WriteLine(drHerTools[0]["magentoid"].ToString() + "; " + drHerTools[0]["qtyonhand"].ToString() + ", dJsonItem.qty: " + dJsonItem.qty + "; BO " + drHerTools[0]["htbo"].ToString() + " : " + dJsonItem.backorders);
                                }
                            }
                        }
                        //else
                        //{
                            //SKU in magento is probably inactive in hertools, don't worry about this yet, do report later
                            //Console.WriteLine(dJsonItem.product_id + " not found");
                        //}
                    }
                }

                //Update values not received in from API to be helpful
                Helper.Sql_Misc_NonQuery("UPDATE Inventory_Issues SET itemid = RII.itemid, SKU = REPLACE(LOWER(herroom.dbo.fn_Magento_SkuID_From_ItemID(RII.Itemid)), '/', '-') "
                    + " FROM Inventory_Issues INV INNER JOIN Herroom..Items RII ON RII.Magentoid = INV.magentoid WHERE resolved = 0");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                bReturnvalue = false;
            }
            return bReturnvalue;
        }


        // ONLY USE TO PROCESS middleware id for now... WIP  
            public static Boolean Process_Low_Inventory(long Middlewareid = 0)
        {
            Boolean bReturnvalue = true;
            DataTable dt;

            try
            {
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, status FROM communications..middleware WHERE id = " + Middlewareid.ToString() + " AND source_table = 'stockItems/lowStock' AND endpoint_method = 'GET'");
                }
                else
                {
                    //INSERT NEW ROW AND USE THAT
                    Helper.Sql_Misc_NonQuery("INSERT communications..middleware(posted, status, worker_id, endpoint_method, endpoint_name, source_id, source_table) "
                    + " SELECT Getdate(), 100, 0, 'GET', 'V1/stockItems/lowStock?scopeId=0&qty=1&pageSize=200000', 'stockItems/lowStock', 'stockItems/lowStock' ");

                    dt = Helper.Sql_Misc_Fetch("SELECT TOP 1 id, status FROM communications..middleware WHERE Status=100 AND source_table = 'stockItems/lowStock' AND endpoint_method = 'GET' ORDER BY ID DESC");
                }
                // takes 10-20 min to run
                if (dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["status"].ToString() == "100")
                    {
                        Helper.MagentoApiPush(long.Parse(dt.Rows[0]["id"].ToString()));
                    }
 
                    Process_Low_Inventory_Report(long.Parse(dt.Rows[0]["id"].ToString()));
                }
            }
            catch (Exception ex)
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine("ERROR Inventory..Process_Low_Inventory() :" + ex.ToString());
                }

                bReturnvalue = false;
            }

            return bReturnvalue;
        }
    }

    class Staging
    {
        public static Boolean Process_Middleware_Staging(long Middlewareid = 0, string SourceTable = "")
        {
            Boolean bReturnvalue = true;
            DataTable dt = new DataTable();
            String APIReturn = "";

            try
            {
                if (Middlewareid > 0)
                {
                    dt = Helper.Sql_Misc_Fetch("SELECT id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE id = " + Middlewareid.ToString());
                }
                else
                {
                    if (SourceTable.Length > 2)
                    {
                        dt = Helper.Sql_Misc_Fetch("SELECT TOP 500 id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE Status=600 AND LoadStaging=100 AND source_table = '" + SourceTable + "' Order By Posted ");
                    }
                    else
                    {
                        dt = Helper.Sql_Misc_Fetch("SELECT TOP 500 id, source_id, endpoint_method, ISNULL(batchid,'') [batchid] FROM Communications..Middleware with (nolock) WHERE Status=600 AND LoadStaging=100 Order By Posted ");
                    }
                }

                Helper.Middleware_Status_Staging_Update(dt, 200);
                
                foreach (DataRow dr in dt.Rows)
                {
                    APIReturn = Helper.MagentoApiPushStaging(long.Parse(dr["id"].ToString())).Replace("\"", "");

                    if (APIReturn.Length > 4 && APIReturn.Substring(0, 5).ToUpper() == "ERROR")
                    {
                        //Middleware should already by set to 500
                        Console.WriteLine("ERROR for Process_Product_Configurable_Middleware Processing: " + dr["source_id"].ToString());
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET LoadStaging=700 WHERE LoadStaging <> 701 AND id = " + dr["id"].ToString());
                    }
                    else  
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE Communications..Middleware SET LoadStaging=600 WHERE ID = " + dr["id"].ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Process_Middleware_Staging(" + Middlewareid.ToString() + ", " + SourceTable + "): " + ex.ToString());
                bReturnvalue = false;
            }

            return bReturnvalue;
        }
    }
}
