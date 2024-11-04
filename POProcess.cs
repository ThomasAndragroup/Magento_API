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
    class POProcess
    {
        //Devmode=2:: don't run, Devmode=1:: run with output, Devmode=0:: run with no output
        public static Boolean PO_HT_Processxx(long Middlewareid = 0)
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

        //Call this 
        public static Boolean PODropShipProcessOrders(string Orderno)
        {
            Boolean  ReturnValue = true ;
            System.Data.DataTable dt;
            string SqlString;

            SqlString = "SELECT DISTINCT OrderNo, RSS.ManufacturerCode FROM Hercust..Orders OO with(nolock) INNER JOIN HerCust..Items II with(nolock) ON Ordernum = Orderno "
                    + " INNER JOIN Herroom..Items RII with(nolock) ON UPC = SKU "
                    + " INNER JOIN Herroom..Styles RSS with(nolock) ON RSS.Stylenumber = RII.Stylenumber "
                    + " LEFT OUTER JOIN Hercust..Orders_DropShipPO DS with(nolock) ON DS.ordernum = OO.orderno AND DS.MfrCode = RSS.ManufacturerCode "
                    + " WHERE DROPSHIP = 1 "
                    + " AND DS.id IS NULL "
                    + " AND ISNULL(CCIntegritycheck, '') <> '' "
                    + " AND Held = 0 AND Cancel = 0 AND Shipped = 0 AND ISNULL(DotComSent,0) = 0 "
                    + " AND OO.Backorder = 0 AND II.Backorder = 0 ";

            if (Orderno != "")
            {
                SqlString += " AND OO.Orderno = " + Orderno;
            }

            dt = Helper.Sql_Misc_Fetch(SqlString);       

            foreach (System.Data.DataRow dr in dt.Rows)
            {
                if (PODropShipCreate(Orderno, dr["ManufacturerCode"].ToString()))
                {
                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("UPDATE Hercust..Orders SET Shipped=1 WHERE Orderno = " + Orderno);
                    }
                    else
                    {
                        Helper.Sql_Misc_NonQuery("UPDATE Hercust..Orders SET Shipped=1 WHERE Orderno = " + Orderno);
                        Helper.Sql_Misc_NonQuery("INSERT Hercust..OrderNotes(Orderno, Posted, Author, Message, isPublic, isAction, Ticklersent) SELECT " + Orderno + ", Getdate(), 'POProcess API', 'PO File Created, Sent to SPS for Vendor: " + dr["ManufacturerCode"].ToString() + "', 0, 0, 0");
                    }
                }
            }

            return ReturnValue;
        }

        //ONLY HAVE MCP Create middleware row AFTER Dropship order invoices, send alerts otherwise !!
        //      it should invoice IMMEDIATELY
        // Process_DropShipPOs(): Select * from Communications..Middleware WHERE Status = 100 AND Source_Table = 'DropshipPO'... source_id is Orderno/po Number
        // Need to rewrite this for Order POs: EXEC HerRoom..[proc_po_process_json] @PONumber = ... to get Hercust..ORders/items QTY where Dropship=1, etc
        // DON'T need: PO_PDFLabels_Send

        //USE THIS FOR DROPSHIP POCreation, only get data from orders..dropship = 1 and use customer data
        // for address, etc, use Orderno as PO#, and update hercust..orders and shipped when PO is sent??
        // only send where orders.ccintegritycheck <> '' 
        // Update herroom..items.VendorQtyOH for each sku
        // DO THIS IN CALLING Funtion!! This func just created the PO and moves it to SPS FTP folder

        //
        // CALLING FUNCTION will have to call PODropshipCreate() for each MFR in DS order
        // Create a PO for each MFR/Order#  so PONumber is OrderNo + MfrCode ???? 
        // ?? Domestic orders only ?? 
        public static Boolean PODropShipCreate(string Orderno, string MfrCode)
        {
            Boolean ReturnValue = true;
            string OutFilename;
            string FilenameXls;

            HelperModels.POOutput POOut;
            HelperModels.POEDI.Address Addr;
            HelperModels.POEDI.Notes note;
            HelperModels.POEDI.Terms term;
            HelperModels.POEDI.ProductorItemDescriptions ProductorItemDescription;
            HelperModels.POEDI.FOBRelatedInstruction FOBRelatedInstr;
            HelperModels.POEDI.LineItem[] LIs;

            DataSet ds = new DataSet();
            System.Data.DataTable dtSku;
            System.Data.DataTable dtCustomer;
            System.Data.DataTable dtCost;
            int SkuRow = 11;

            //SPS850DROPSHIPFILEOUTPUT

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

                OutFilename = ConfigurationManager.AppSettings["SPS850DROPSHIPFILEOUTPUT"] + "PO" + Orderno + ".json";
                FilenameXls = ConfigurationManager.AppSettings["SPS850DROPSHIPFILEOUTPUT"] + "PO" + Orderno + ".xlsx";

                //Datatables: 1-Sku info; 2-Mfr info
                ds = Helper.Sql_Misc_Fetch_Dataset("EXEC HerRoom..[proc_po_process_dropship_json] @OrderNo = '" + Orderno + "', @MFRCode = '" + MfrCode + "'");



                if (MagetnoProductAPI.DevMode > 1)
                {
                    Console.WriteLine("DS: " + ds.Tables.Count.ToString());
                }

                if (ds.Tables.Count >= 3)
                {
                    dtSku = ds.Tables[0];
                    dtCustomer = ds.Tables[1];
                    dtCost = ds.Tables[2];

                    xlWorkSheet.Name = "PO " + Orderno;

                    WriteXLSHeader(xlWorkSheet, dtCustomer.Rows[0]);

                    ////////////////////////////////////////////////////////// 
                    HelperModels.POEDI.PO htPO = new HelperModels.POEDI.PO();
                    HelperModels.POEDI.Header htHeader = new HelperModels.POEDI.Header();
                    HelperModels.POEDI.OrderHeader htOrderheader = new HelperModels.POEDI.OrderHeader();
                    htPO.header = new HelperModels.POEDI.Header();
                    htPO.header.OrderHeader = new HelperModels.POEDI.OrderHeader();
                    htPO.header.OrderHeader.PurchaseOrderNumber = Orderno;
                    htPO.header.OrderHeader.TsetPurposeCode = "00";
                    htPO.header.OrderHeader.PrimaryPOTypeCode = "SA";
                    htPO.header.OrderHeader.PurchaseOrderDate = DateTime.Parse(dtCustomer.Rows[0]["orderdate"].ToString());
                    htPO.header.OrderHeader.Vendor = dtCustomer.Rows[0]["manufacturercode"].ToString();

                    htPO.header.Date = new List<HelperModels.POEDI.DateInfo>();
                    HelperModels.POEDI.DateInfo dateinfo010 = new HelperModels.POEDI.DateInfo();
                    dateinfo010.Datetimequalifier = "010";
                    dateinfo010.Date = DateTime.Parse(dtCustomer.Rows[0]["POStartShipDate1"].ToString()); // ????
                    htPO.header.Date.Add(dateinfo010);

                    HelperModels.POEDI.DateInfo dateinfo001 = new HelperModels.POEDI.DateInfo();
                    dateinfo001.Datetimequalifier = "001";
                    dateinfo001.Date = DateTime.Parse(dtCustomer.Rows[0]["POCancelDate1"].ToString()); // ?????
                    htPO.header.Date.Add(dateinfo001);

                    htPO.header.address = new List<HelperModels.POEDI.Address>();
                    Addr = new HelperModels.POEDI.Address();
                    Addr.AddressTypeCode = "ST";
                    Addr.AddressLocationNumber = dtCustomer.Rows[0]["POAddressLocationNumber"].ToString();
                    Addr.LocationCodeQualifier = "92;";
                    Addr.Address1 = dtCustomer.Rows[0]["shipaddr"].ToString();
                    Addr.City = dtCustomer.Rows[0]["shipcity"].ToString();
                    Addr.State = dtCustomer.Rows[0]["shipst"].ToString();
                    Addr.PostalCode = dtCustomer.Rows[0]["shipzip"].ToString();

                    htPO.header.address.Add(Addr);

                    htPO.header.Terms = new List<HelperModels.POEDI.Terms>();
                    term = new HelperModels.POEDI.Terms();
                    term.termsDescription = dtCustomer.Rows[0]["POTerms"].ToString();
                    htPO.header.Terms.Add(term);

                    htPO.ordersummary = new HelperModels.POEDI.OrderSummary();
                    htPO.ordersummary.TotalAmount = dtCost.Rows[0]["cost"].ToString();
                    htPO.ordersummary.TotalLineItemNumber = dtCost.Rows[0]["rows"].ToString();
                    htPO.ordersummary.TotalQuantity = dtCost.Rows[0]["quantity"].ToString();

                    /////////////////////////////////////////////////
                    List<HelperModels.POEDI.Item> htItems = new List<HelperModels.POEDI.Item>();
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
                        htLineItem.orderline.OrderQty = dr["OrderQty"].ToString();
                        htLineItem.orderline.OrderQtyUOM = "EA";
                        htLineItem.orderline.PurchasePrice = dr["ourcost"].ToString();
                        htLineItem.orderline.Color = dr["colorname"].ToString();
                        htLineItem.orderline.Size = dr["size"].ToString();

                        //if (IncludeOptionalCode)
                        //{
                        //    htLineItem.orderline.ExtendedItemTotal = dr["extendedcost"].ToString();
                        //}

                        htLineItem.productorItemDescriptions = new List<HelperModels.POEDI.ProductorItemDescriptions>();
                        ProductorItemDescription = new HelperModels.POEDI.ProductorItemDescriptions();
                        ProductorItemDescription.ProductCharacteristicCode = "08";  // dr["postyle"].ToString();
                        ProductorItemDescription.ProductDescription = dr["productname"].ToString();
                        htLineItem.productorItemDescriptions.Add(ProductorItemDescription);

                        htLineItems.Add(htLineItem);
                        SkuRow++;

                        POOut = new HelperModels.POOutput();

                        POOut.LineNumber = (SkuRow - 11).ToString();
                        POOut.Manufacturer = dtCustomer.Rows[0]["manufacturername"].ToString();
                        POOut.PONumber = "D" + Orderno;
                        POOut.Posted = dtCustomer.Rows[0]["startdate"].ToString();
                        POOut.Style = dr["stylenumber"].ToString();
                        POOut.Description = dr["productname"].ToString();
                        POOut.Color = dr["colorname"].ToString();
                        POOut.ColorCode = dr["colorcode"].ToString();
                        POOut.Size = dr["size"].ToString();
                        POOut.UPC = dr["upc"].ToString();
                        POOut.QtyOrdered = dr["OrderQty"].ToString();
                        POOut.Cost = "$" + dr["ourcost"].ToString();
                        POOut.ExtCost = "$" + dr["extendedcost"].ToString();
                        POOut.Receive = "in process";
                        POOut.ColorOverride = dr["pocoloroverride"].ToString();
                        POOut.StyleOverride = dr["postyleoverride"].ToString();
                        POOut.Closeout = dr["closeout"].ToString();
                        POOut.SKUCloseout = dr["upccloseout"].ToString();
                        POOut.Backorder = dr["bo"].ToString();

                        WriteXLSLine(xlWorkSheet, POOut, SkuRow - 1);
                    }


                    // XLS FOOTER INFO /////////////////////////////////////////////////////////
                    WriteXLSFooter(xlWorkSheet, dtCustomer.Rows[0], dtCost.Rows[0], SkuRow);

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

                    foreach (DataRow dr in dtSku.Rows)
                    {
                        if (MagetnoProductAPI.DevMode > 1)
                        {
                            Console.WriteLine("INSERT Hercust..Orders_DropshipPO(ordernum, mfrcode, itemnum, podate, posent, ASN_Received, SKU) SELECT " + Orderno + ", '" + MfrCode + "', " + dr["itemnum"].ToString() + " Getdate(), 1, 0, '" + dr["UPC"].ToString() + "; ");

                            Console.WriteLine("INSERT Hercust..POArchive(UPC, NumberExp, DateIn, UnitCost, PONumber QtyRcvd) "
                                + " VALUES('" + dr["upc"].ToString() + "', " + dr["OrderQty"].ToString() + ", GetDate(), " + dr["ourcost"].ToString() + ", 'D'" + Orderno + "', 0)");
                        }
                        else
                        {
                            Helper.Sql_Misc_NonQuery("INSERT Hercust..Orders_DropshipPO(ordernum, mfrcode, itemnum, podate, posent, ASN_Received) SELECT " + Orderno + ", '" + MfrCode + "', " + dr["itemnum"].ToString() + ", Getdate(), 1, 0 ");

                            Helper.Sql_Misc_NonQuery("INSERT Hercust..POArchive(UPC, NumberExp, DateIn, UnitCost, PONumber QtyRcvd) "
                                + " VALUES('" + dr["upc"].ToString() + "', " + dr["OrderQty"].ToString() + ", GetDate(), " + dr["ourcost"].ToString() + ", 'D'" + Orderno + "', 0)");
                        }
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

            }
            catch (Exception ex)
            {
                if (MagetnoProductAPI.DevMode > 0 || MagetnoProductAPI.DevMode == -1)
                {
                    Console.WriteLine("PO XLS ERROR: " + ex.ToString());
                }
                ReturnValue = false;
            }

            return ReturnValue;
        }

        public static Boolean POCreate(string MFRCode, string PONumber, Boolean IncludeOptionalCode = false)
        {
            Boolean ReturnValue = true;
            string OutFilename;
            string FilenameXls;
            //string OutputLine;
            
            HelperModels.POOutput POOut;
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

                    WriteXLSHeader(xlWorkSheet, dtMfr.Rows[0]);

                    //////////////////////////////////////////////////////////            
                    HelperModels.POEDI.PO htPO = new HelperModels.POEDI.PO();
                    HelperModels.POEDI.Header htHeader = new HelperModels.POEDI.Header();
                    HelperModels.POEDI.OrderHeader htOrderheader = new HelperModels.POEDI.OrderHeader();
                    htPO.header = new HelperModels.POEDI.Header();
                    htPO.header.OrderHeader = new HelperModels.POEDI.OrderHeader();
                    htPO.header.OrderHeader.PurchaseOrderNumber = dtMfr.Rows[0]["ponumber"].ToString();
                    htPO.header.OrderHeader.TsetPurposeCode = "00";
                    htPO.header.OrderHeader.PrimaryPOTypeCode = "SA";
                    htPO.header.OrderHeader.PurchaseOrderDate = DateTime.Parse(dtMfr.Rows[0]["startdate"].ToString());
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

                        htLineItem.productorItemDescriptions = new List<HelperModels.POEDI.ProductorItemDescriptions>();
                        ProductorItemDescription = new HelperModels.POEDI.ProductorItemDescriptions();
                        ProductorItemDescription.ProductCharacteristicCode = "08";  // dr["postyle"].ToString();
                        ProductorItemDescription.ProductDescription = dr["productname"].ToString();
                        htLineItem.productorItemDescriptions.Add(ProductorItemDescription);

                        htLineItems.Add(htLineItem);
                        SkuRow++;
                      
                        POOut = new HelperModels.POOutput();

                        POOut.LineNumber = (SkuRow - 11).ToString();
                        POOut.Manufacturer = dtMfr.Rows[0]["manufacturername"].ToString();
                        POOut.PONumber = PONumber;
                        POOut.Posted = dtMfr.Rows[0]["startdate"].ToString();
                        POOut.Style = dr["stylenumber"].ToString();
                        POOut.Description = dr["productname"].ToString();
                        POOut.Color = dr["colorname"].ToString();
                        POOut.ColorCode = dr["colorcode"].ToString();
                        POOut.Size = dr["size"].ToString();
                        POOut.UPC = dr["upc"].ToString();
                        POOut.QtyOrdered = dr["poorder"].ToString();
                        POOut.Cost = "$" + dr["ourcost"].ToString();
                        POOut.ExtCost = "$" + dr["extendedcost"].ToString();
                        POOut.Receive = "in process";
                        POOut.ColorOverride = dr["pocoloroverride"].ToString();
                        POOut.StyleOverride = dr["postyleoverride"].ToString();
                        POOut.Closeout = dr["closeout"].ToString();
                        POOut.SKUCloseout = dr["upccloseout"].ToString();
                        POOut.Backorder = dr["bo"].ToString();

                        WriteXLSLine(xlWorkSheet, POOut, SkuRow - 1);
                    }


                    // XLS FOOTER INFO /////////////////////////////////////////////////////////
                    WriteXLSFooter(xlWorkSheet, dtMfr.Rows[0], dtCost.Rows[0], SkuRow);   

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
                                + " SELECT '" + dt.Rows[0]["poemail"].ToString() + "','buyers@andragroup.com','thomas@andragroup.com', '" + EmailBody + "', 'PO Shipping Labels for PO:" + PONumber + "', '" + fi.FullName + "', Getdate(), 0");
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
   
        public static Boolean Incoming_856File_Process(Boolean SkipFileMove = false)
        {
            Boolean ReturnValue = true;
            DirectoryInfo di;
            DirectoryInfo diProcess;
            string FileFolder;
            String DirectoryArchive;
            String DirectoryError;
            String DirectoryFTP;
            //System.Data.DataTable dt;
            String PONumber;
            string DirectoryDaily;
            Boolean FileMoveSuccess;

            //find all files, process the unprocessed ones;
            FileFolder = ConfigurationManager.AppSettings["SPS856FILEINCOMING"];
           
            DirectoryError = ConfigurationManager.AppSettings["SPS856FILEINCOMINGERROR"];
            DirectoryFTP = ConfigurationManager.AppSettings["SPS856FILEINCOMINGFTP"];
            DirectoryArchive = ConfigurationManager.AppSettings["SPS856FILEINCOMINGARCHIVE"];
            DirectoryDaily = ConfigurationManager.AppSettings["SPS856FILEINCOMINGDAILYFILES"];

            di = new DirectoryInfo(DirectoryFTP);
            FileInfo[] Files = di.GetFiles();

            if (SkipFileMove)
            {
                FileMoveSuccess = true;
            }
            else
            {
                FileMoveSuccess = Incoming_856File_FTP_Move();
            }

            if (FileMoveSuccess)
            {
                diProcess = new DirectoryInfo(FileFolder);
                FileInfo[] FilesProcess = diProcess.GetFiles();

                foreach (FileInfo fi in FilesProcess)
                {
                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine(" !! PROCESS File: " + fi.FullName);
                    }

                    Helper.Sql_Misc_NonQuery("INSERT hercust..po_856_incoming_file([Filename], processed, datestamp) SELECT '" + fi.Name + "', 100, Getdate()  WHERE 0 = (SELECT COUNT(*) FROM hercust..po_856_incoming_file WHERE [Filename] = '" + fi.FullName + "') ");

                    if (Incoming_856File(fi.FullName, fi.Name, out PONumber))
                    {
                        if (PONumber.Substring(0, 1).ToUpper() == "D")
                        {
                            //DROPSHIP PO ASN FOLDER !!!! 
                            //fi.MoveTo(DirectoryArchive + fi.Name);
                            // AUTOMATE SHOULD EMAIL THESE NORMALLY WITH OTHER ASN xls ???

                            Helper.Sql_Misc_NonQuery("INSERT hercust..orderNotes() WHERE Orderno = REPLACE('" + PONumber + "', 'D', '')");
                        }
                        else
                        {
                            fi.MoveTo(DirectoryArchive + fi.Name);
                        }

                    }
                    else
                    {
                        fi.MoveTo(DirectoryError + fi.Name);
                    }
                }
            }

            return ReturnValue;
        }

        //2024-10-24
        //Records what should have been sent to SPS via FTP (some files aren't making it)
        public static Boolean Incoming_850_FTP_Record(string ProcessName = "Automate: FTP Outbound Files")
        {
            Boolean ReturnValue = true;
            DirectoryInfo di;
            string FileFolder;
            String PONumber;
            System.Data.DataTable dt;
            string DirectoryDaily;

            //find all files, process the unprocessed ones;
            FileFolder = ConfigurationManager.AppSettings["SPS850FILEOUTPUT"];

            di = new DirectoryInfo(FileFolder);
            FileInfo[] Files = di.GetFiles();

            //Move Files to be Processed to Root Directory
            foreach (FileInfo fi in Files)
            {
                if (MagetnoProductAPI.DevMode > 1)
                {
                    Console.WriteLine("File: " + fi.FullName); // + "; " + fi.FullName.Replace(@"\FTP", ""));
                    Console.WriteLine("File: " + fi.Name);
                    Console.WriteLine(" ------------- ");
                }

                if (fi.Name.EndsWith(".json"))
                {
                    PONumber = fi.Name.Replace(".json", "");
                        
                    if (MagetnoProductAPI.DevMode > 0)
                    {
                        Console.WriteLine("TO RECORD File: " + fi.FullName + "; PO: " + PONumber);
                        Console.WriteLine("INSERT Herroom..PO_FTP_Log(PONumber, POFilename, FTPSystem, Datestamp, Verified_Xfer) "
                            + " SELECT '" + PONumber + "', '" + fi.FullName + "', '" + ProcessName + "', Getdate(), 0");

                    }
                    if (MagetnoProductAPI.DevMode < 2)
                    {
                        Helper.Sql_Misc_NonQuery("INSERT Herroom..PO_FTP_Log(PONumber, POFilename, FTPSystem, Datestamp, Verified_Xfer) "
                            + " SELECT '" + PONumber + "', '" + fi.FullName + "', 'Automate: FTP Outbound Files', Getdate(), 0");

                    }
                }
            }

            return ReturnValue;
        }

        //2024-10-24 !!
        public static Boolean SPS_Report_Reader(string FileName)
        {
            Boolean ReturnValue = true;
            string xText;
            FileName = @"\\task\Automate\Working\PO-SPS\REPORTS\DailyReceived.csv";
            String[] LineInfo;
            
            StreamReader sr = new StreamReader(FileName);
            //xText = sr.ReadToEnd();
            //Console.WriteLine(xText);

            while (!sr.EndOfStream)
            {
                //strContent.Add(reader.ReadLine());
                xText = sr.ReadLine();
               
                xText = xText.Replace(@"""", "");

                if (xText.Substring( 0, 3) == "850")
                {
                    LineInfo = xText.Split(Char.Parse(","));
                    //Console.WriteLine(xText);

                    if (LineInfo.Length >= 6)
                    {
                        Console.WriteLine(LineInfo[0] + "; " + LineInfo[3] + "; " + LineInfo[5] + "; " + LineInfo[6]);

                        if (MagetnoProductAPI.DevMode > 0)
                        {
                            Console.WriteLine("INSERT Communications..SPS_DailyReceived([SPSFile], PONumber, RecStatus, DateReceived, DateStamp, Processed) "
                            + "SELECT '" + FileName + "', '" + LineInfo[6] + "', '" + LineInfo[5] + "', '" + LineInfo[3] + "', Getdate(), 0 "
                            + "WHERE 0 = (SELECT Count(*) FROM Communications..SPS_DailyReceived WHERE POnumber = '' AND DateReceived = '')");
                        }

                        Helper.Sql_Misc_NonQuery("INSERT Communications..SPS_DailyReceived([SPSFile], PONumber, RecStatus, DateReceived, DateStamp, Processed) "
                            + "SELECT '" + FileName + "', '" + LineInfo[6] + "', '" + LineInfo[5] + "', '" + LineInfo[3] + "', Getdate(), 0 "
                            + "WHERE 0 = (SELECT Count(*) FROM Communications..SPS_DailyReceived WHERE POnumber = '" + LineInfo[6] + "' AND DateReceived = '" + LineInfo[3] + "')");

                    }
                }
            }

            //Update Tables
            Helper.Sql_Misc_NonQuery("UPDATE Herroom..PO_FTP_Log SET Verified_Xfer = 1 "
                + " FROM Herroom..PO_FTP_Log POT "
                + " LEFT OUTER JOIN Communications..SPS_DailyReceived SPS ON REPLACE(SPS.PONumber, 'PO', '') = REPLACE(POT.PONumber, 'PO', ''); "
                + " UPDATE Communications..SPS_DailyReceived SET Processed = 1 "
                + " FROM Communications..SPS_DailyReceived SPS "
                + " INNER JOIN Herroom..PO_FTP_Log POT ON REPLACE(SPS.PONumber, 'PO', '') = REPLACE(POT.PONumber, 'PO', '')");


            return ReturnValue;
        }


        public static Boolean Incoming_856File_FTP_Move()
        { 
            Boolean ReturnValue = true;
            DirectoryInfo di;
            string FileFolder;
            String DirectoryArchive;
            String DirectoryError;
            String DirectoryFTP;
            System.Data.DataTable dt;
            string DirectoryDaily;

            //find all files, process the unprocessed ones;
            FileFolder = ConfigurationManager.AppSettings["SPS856FILEINCOMING"];

            DirectoryError = ConfigurationManager.AppSettings["SPS856FILEINCOMINGERROR"];
            DirectoryFTP = ConfigurationManager.AppSettings["SPS856FILEINCOMINGFTP"];
            DirectoryArchive = ConfigurationManager.AppSettings["SPS856FILEINCOMINGARCHIVE"];
            DirectoryDaily = ConfigurationManager.AppSettings["SPS856FILEINCOMINGDAILYFILES"];

            di = new DirectoryInfo(DirectoryFTP);
            FileInfo[] Files = di.GetFiles();

            //Move Files to be Processed to Root Directory
            foreach (FileInfo fi in Files)
            {
                if (MagetnoProductAPI.DevMode > 0)
                {
                    Console.WriteLine("File: " + fi.FullName + "; " + fi.FullName.Replace(@"\FTP", ""));
                    Console.WriteLine("File: " + fi.Name);
                    Console.WriteLine(" ------------- ");
                }

                // USE FILENAME ONLY WITHOUT PATH, TOO CONFUSING AND UNNECCESSARY 
                dt = new System.Data.DataTable();  // clear table 
                dt = Helper.Sql_Misc_Fetch("SELECT COUNT(*) [cnt] FROM hercust..po_856_incoming_file WHERE Filename = '" + fi.Name + "'");
                if (dt.Rows.Count > 0 && dt.Rows[0]["cnt"].ToString() == "0")
                {
                    //Move File to Main folder to process
                    fi.MoveTo(FileFolder + fi.Name);
                    // 2024-09-17 for ASN_Daily_Report_Process() instead, keep daily one-PO files but email the compound daily one
                    //fi.MoveTo(DirectoryDaily + fi.Name);
                }
            }

            return ReturnValue;
        }

        public static Boolean Incoming_856File(string FilePath, string ShortFileName, out string PONumber)
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

            PONumber = "";

            //FileArchive = ConfigurationManager.AppSettings["SPS856FILEARCHIVE"];
            FileArchive = FilePath.Replace(@"/inbound/", @"/inbound/archive/");

            //Set to status 200 while processing, then 600, 700
            Helper.Sql_Misc_NonQuery("UPDATE HerCust..[PO_856_Incoming_file] SET Processed = 200 WHERE Filename = '" + ShortFileName + "';");

            FilenameXls = FilePath.Replace(".json", ".xlsx").Replace(".txt", ".xlsx");
            //2024-09-17: store single-PO files to \DailyFiles so only multi-PO file gets emailed
            FilenameXls = FilenameXls.ToLower().Replace(@"/inbound/", @"/inbound/dailyfiles/").Replace(@"\inbound\", @"\inbound\dailyfiles\");

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

            //Header
            Lineout = "TransactionType,PONum,Buyer,ManufacturerName,PODate,PO Start Ship Date, PO Cancel Date,Stylenumber,Item Description,Color,Color Code,Size,UPC,OurCost,QuantityShipped,QuantityOrdered,BO,ImportedDate,AccountingId,TrackingNumber,ScheduledDelivery,Vendor ShipDate,ShipVia,NumOfCartonsShipped,ShipFromName,ShipFromAddressLineOne, ShipFromAddressLineTwo,ShipFromCity,ShipFromState,ShipFromZip,ShipFromCountry,ShipFromAddressCode,VendorNum,DCCode,TransportationMethod,Status,UOMOfUPCs,POcost,LineNum,Notes,CarrierProNumber,CurrentScheduledDeliveryDate,CarrierRouting";

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
                    PONumber = OL.OrderHeader.PurchaseOrderNumber;

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
                                    + " , (SELECT TOP 1 LineNum FROM HerCust..Expected EX WHERE EX.PONumber = PA.PONumber AND Ex.upc = PA.UPC ORDER BY EX.Invkey DESC) [LineNum]"  
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

                                    Lineout += "," + drPOA["NumberExp"].ToString() + "," + drPOA["bo"].ToString() + "," + drPOA["datein"].ToString() + "," + drPOA["ManufacturerCode"].ToString() + ",'" + SP.Header.ShipmentHeader.BillOfLadingNumber + ", " + SP.Header.ShipmentHeader.CurrentScheduledDeliveryDate + ", " + SP.Header.ShipmentHeader.ShipDate + ",";

                                    Lineout += "," + CartonsinPO + "," + Address_Name;   //OL.OrderHeader.Vendor;

                                    Lineout += "," + Address_Address1 + ",," + Address_City + ",," + Address_Postcode + ",,,,01," + Carrier + ",,Each,$" + drPOA["pocost"].ToString();

                                    Lineout += "," + drPOA["LineNum"].ToString();       // LineNum (before Notes)

                                    //2024-10-03 Per Buyers: CarrierProNumber,CurrentScheduledDeliveryDate,CarrierRouting
                                    Lineout += "," + SP.Header.ShipmentHeader.CarrierProNumber + "," + SP.Header.ShipmentHeader.CurrentScheduledDeliveryDate + ",";

                                    if (SP.Header.CarrierInformation.Count > 0 && SP.Header.CarrierInformation[0].CarrierRouting != null)
                                    {
                                        Lineout += SP.Header.CarrierInformation[0].CarrierRouting;
                                    }
                                    else
                                    {
                                        Lineout += "unknown";
                                    }

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
                                        //Console.WriteLine("INSERT Hercust..PO_856_Incoming(PONumber, FileName, ShippingSerialID, UPC, QtyIn, Notes, Datestamp) SELECT '" + OL.OrderHeader.PurchaseOrderNumber + "', '" + ShortFileName + "', '" + PL.Pack.ShippingSerialID + "', '" + IL.ShipmentLine.ConsumerPackageCode + "', " + IL.ShipmentLine.ShipQty.ToString() + ", '', '" + SP.Header.ShipmentHeader.ShipDate + "' WHERE 0 = (SELECT COUNT(*) FROM Hercust..PO_856_Incoming WHERE PONumber = '" + OL.OrderHeader.PurchaseOrderNumber + "' AND ShippingSerialId = '" + PL.Pack.ShippingSerialID + "' AND UPC = '" + IL.ShipmentLine.ConsumerPackageCode  + "')");
                                        Console.WriteLine("INSERT Hercust..PO_856_Incoming(PONumber, FileName, ShippingSerialID, UPC, QtyIn, Notes, Datestamp, BillOfLadingNumber) SELECT '" + OL.OrderHeader.PurchaseOrderNumber + "', '" + ShortFileName + "', '" + PL.Pack.ShippingSerialID + "', '" + IL.ShipmentLine.ConsumerPackageCode + "', " + IL.ShipmentLine.ShipQty.ToString() + ", '', '" + SP.Header.ShipmentHeader.ShipDate + "' WHERE 0 = (SELECT COUNT(*) FROM Hercust..PO_856_Incoming WHERE PONumber = '" + OL.OrderHeader.PurchaseOrderNumber + "' AND ShippingSerialId = '" + PL.Pack.ShippingSerialID + "' AND UPC = '" + IL.ShipmentLine.ConsumerPackageCode + "', '" + SP.Header.ShipmentHeader.BillOfLadingNumber + "')");
                                        Console.WriteLine("------");
                                    }

                                    //Helper.Sql_Misc_NonQuery("INSERT Hercust..PO_856_Incoming(PONumber, FileName, ShippingSerialID, UPC, QtyIn, Notes, Datestamp) SELECT '" + OL.OrderHeader.PurchaseOrderNumber + "', '" + ShortFileName + "', '" + PL.Pack.ShippingSerialID + "', '" + IL.ShipmentLine.ConsumerPackageCode + "', " + IL.ShipmentLine.ShipQty.ToString() + ", '', '" + SP.Header.ShipmentHeader.ShipDate + "' WHERE 0 = (SELECT COUNT(*) FROM Hercust..PO_856_Incoming WHERE PONumber = '" + OL.OrderHeader.PurchaseOrderNumber + "' AND ShippingSerialId = '" + PL.Pack.ShippingSerialID + "' AND UPC = '" + IL.ShipmentLine.ConsumerPackageCode + "')");
                                    Helper.Sql_Misc_NonQuery("INSERT Hercust..PO_856_Incoming(PONumber, FileName, ShippingSerialID, UPC, QtyIn, Notes, Datestamp, BillOfLadingNumber, CSVLine) SELECT '" + OL.OrderHeader.PurchaseOrderNumber + "', '" + ShortFileName + "', '" + PL.Pack.ShippingSerialID + "', '" + IL.ShipmentLine.ConsumerPackageCode + "', " + IL.ShipmentLine.ShipQty.ToString() + ", '', '" + SP.Header.ShipmentHeader.ShipDate + "', '" + SP.Header.ShipmentHeader.BillOfLadingNumber + "', '" + Lineout.Replace("'", "''") + "' WHERE 0 = (SELECT COUNT(*) FROM Hercust..PO_856_Incoming WHERE PONumber = '" + OL.OrderHeader.PurchaseOrderNumber + "' AND ShippingSerialId = '" + PL.Pack.ShippingSerialID + "' AND UPC = '" + IL.ShipmentLine.ConsumerPackageCode + "')");                                  

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

                                    //2024-09-12 for daily bulk report 
                                    //2024-09-18 moved 
                                   // Helper.Sql_Misc_NonQuery("UPDATE Hercust..PO_856_Incoming SET CSVLine = '" + Lineout.Replace("'", "''") + "' WHERE PONumber = '" + OL.OrderHeader.PurchaseOrderNumber + "' AND UPC = '" + IL.ShipmentLine.ConsumerPackageCode  + "'");
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

                                        //2024-10-03 Per Buyers: CarrierProNumber,CurrentScheduledDeliveryDate,CarrierRouting
                                        Lineout += "," + SP.Header.ShipmentHeader.CarrierProNumber + "," + SP.Header.ShipmentHeader.CurrentScheduledDeliveryDate + ",";

                                        if (SP.Header.CarrierInformation.Count > 0 && SP.Header.CarrierInformation[0].CarrierRouting != null)
                                        {
                                            Lineout += SP.Header.CarrierInformation[0].CarrierRouting;
                                        }
                                        else
                                        {
                                            Lineout += "unknown";
                                        }

                                        if (MagetnoProductAPI.DevMode > 0)
                                        {
                                            Console.WriteLine(Lineout);
                                        }

                                        WriteXLSLine(xlWorkSheet, Lineout, Rowcount, true, "A");
                                        Rowcount++;

                                        if (MagetnoProductAPI.DevMode > 0)
                                        {
                                            Console.WriteLine(Lineout);

                                            //Dropship 
                                            if (PONumber.Substring(0, 1).ToUpper() == "D")
                                            {
                                                Console.WriteLine("UPDATE Hercust..Orders_DropshipPO.ASN_Received = 1 WHERE Ordernum = Orderno AND SKU = '" + IL.ShipmentLine.ConsumerPackageCode + "'");
                                                Console.WriteLine("UPDATE Hercust..POArchive SET QTYRcvd = (ISNULL(QtyRcvd,0) + " + IL.ShipmentLine.ShipQty + ") WHERE PONumber = '' AND UPC = '" + IL.ShipmentLine.ConsumerPackageCode + "'"); 
                                            }

                                            Console.WriteLine(" ----------------- ");
                                            Console.WriteLine("INSERT Hercust..PO_856_Incoming(PONumber, FileName, ShippingSerialID, UPC, QtyIn, Notes, Datestamp, BillOfLadingNumber, CSVLine) SELECT '" + OL.OrderHeader.PurchaseOrderNumber + "', '" + ShortFileName + "', '" + PL.Pack.ShippingSerialID + "', '" + IL.ShipmentLine.ConsumerPackageCode + "', " + IL.ShipmentLine.ShipQty.ToString() + ", 'SKU NOT FOUND IN: " + IL.ShipmentLine.ConsumerPackageCode + ": " + IL.ProductOrItemDescription[0].ProductDescription + "', '" + SP.Header.ShipmentHeader.ShipDate + "', '" + SP.Header.ShipmentHeader.BillOfLadingNumber + "', '" + Lineout.Replace("'", "''") + "' WHERE 0 = (SELECT COUNT(*) FROM Hercust..PO_856_Incoming WHERE PONumber = '" + OL.OrderHeader.PurchaseOrderNumber + "' AND ShippingSerialId = '" + PL.Pack.ShippingSerialID + "' AND UPC = '" + SP.Header.ShipmentHeader.BillOfLadingNumber + "')");
                                            Console.WriteLine(" ----------------- ");
                                        }

                                        Helper.Sql_Misc_NonQuery("INSERT Hercust..PO_856_Incoming(PONumber, FileName, ShippingSerialID, UPC, QtyIn, Notes, Datestamp, BillOfLadingNumber, CSVLine) SELECT '" + OL.OrderHeader.PurchaseOrderNumber + "', '" + ShortFileName + "', '" + PL.Pack.ShippingSerialID + "', '" + IL.ShipmentLine.ConsumerPackageCode + "', " + IL.ShipmentLine.ShipQty.ToString() + ", 'SKU NOT FOUND IN: " + IL.ShipmentLine.ConsumerPackageCode + ": " + IL.ProductOrItemDescription[0].ProductDescription + "', '" + SP.Header.ShipmentHeader.ShipDate + "', '" + SP.Header.ShipmentHeader.BillOfLadingNumber + "', '" + Lineout.Replace("'", "''") + "' WHERE 0 = (SELECT COUNT(*) FROM Hercust..PO_856_Incoming WHERE PONumber = '" + OL.OrderHeader.PurchaseOrderNumber + "' AND ShippingSerialId = '" + PL.Pack.ShippingSerialID + "' AND UPC = '" + SP.Header.ShipmentHeader.BillOfLadingNumber + "')");

                                        // !!!!!!!!!!!!!
                                        if (PONumber.Substring(0, 1).ToUpper() == "D")
                                        {
                                            Helper.Sql_Misc_NonQuery("UPDATE Hercust..Orders_DropshipPO.ASN_Received = 1 WHERE Ordernum = Orderno AND SKU = '" + IL.ShipmentLine.ConsumerPackageCode + "'");
                                            Helper.Sql_Misc_NonQuery("UPDATE Hercust..POArchive SET QTYRcvd = (ISNULL(QtyRcvd,0) + " + IL.ShipmentLine.ShipQty + ") WHERE PONumber = '' AND UPC = '" + IL.ShipmentLine.ConsumerPackageCode + "'" );
                                        }

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

        //Process previous days entries: Hercust..PO_856_Incoming 
        public static Boolean ASN_Daily_Report_Process() //string DateToProcess = "")
        {
            Boolean ReturnValue = true;
            System.Data.DataTable dt;
            String Lineout;
            String FilenameXls;
            string FileFolder;
            String DirectoryArchive;
            String DirectoryError;
            String DirectoryFTP;
            int Rowcount = 2;

            FileFolder = ConfigurationManager.AppSettings["SPS856FILEINCOMING"];
            DirectoryArchive = ConfigurationManager.AppSettings["SPS856FILEINCOMINGARCHIVE"];
            DirectoryError = ConfigurationManager.AppSettings["SPS856FILEINCOMINGERROR"];
            DirectoryFTP = ConfigurationManager.AppSettings["SPS856FILEINCOMINGFTP"];

           // DateTime now = DateTime.Now;

            FilenameXls = FileFolder + "ASNReport_" + DateTime.Now.ToString("yyyy_MM_dd_HHmm") + ".xlsx";

            /*
            if (DateToProcess == "")
            {
                dt = Helper.Sql_Misc_Fetch("SELECT ponumber, upc, ISNULL(CSVLine,'') [CSVLine] FROM Hercust..PO_856_Incoming WHERE CONVERT(Date, datestamp) = CONVERT(Date, Dateadd(day,-1, Getdate())) ORDER BY PONumber, UPC  ");
            }
            else  // work on date ranges later !!! 
            {
                dt = Helper.Sql_Misc_Fetch("SELECT ponumber, upc, ISNULL(CSVLine,'') [CSVLine] FROM Hercust..PO_856_Incoming WHERE CONVERT(Date, datestamp) = CONVERT(Date, '" + DateToProcess + "') ORDER BY PONumber, UPC  ");
            }
            */
            //2024-10-02 REPORT SHOULD HAVE ALL OPEN POS (Expected)
            //dt = Helper.Sql_Misc_Fetch("SELECT INC.ponumber, INC.upc, ISNULL(CSVLine,'') [CSVLine] FROM Hercust..Expected EX INNER JOIN Hercust..PO_856_Incoming INC ON INC.PONumber = EX.PONumber AND INC.UPC = EX.upc AND ISNULL(csvLine, '') <> '' ORDER BY 1, 2; ");

            dt = Helper.Sql_Misc_Fetch("SELECT INC.ponumber, INC.upc, ISNULL(CSVLine, '') [CSVLine] " 
                    + " FROM Hercust..PO_856_Incoming INC " 
                    + " WHERE PONumber IN(SELECT DISTINCT PONumber FROM Hercust..Expected) AND ISNULL(csvLine, '') <> '' ORDER BY 1, 2;" );


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

            //Header
            Lineout = "TransactionType,PONum,Buyer,ManufacturerName,PODate,PO Start Ship Date, PO Cancel Date,Stylenumber,Item Description,Color,Color Code,Size,UPC,OurCost,QuantityShipped,QuantityOrdered,BO,ImportedDate,AccountingId,TrackingNumber,ScheduledDelivery,Vendor ShipDate,ShipVia,NumOfCartonsShipped,ShipFromName,ShipFromAddressLineOne, ShipFromAddressLineTwo,ShipFromCity,ShipFromState,ShipFromZip,ShipFromCountry,ShipFromAddressCode,VendorNum,DCCode,TransportationMethod,Status,UOMOfUPCs,POcost,LineNum,Notes,CarrierProNumber,CurrentScheduledDeliveryDate,CarrierRouting";

            WriteXLSLine(xlWorkSheet, Lineout, 1);

            foreach (DataRow dr in dt.Rows)
            {
                if (dr["CSVLine"].ToString() != "")
                {
                    WriteXLSLine(xlWorkSheet, dr["CSVLine"].ToString(), Rowcount);
                    Rowcount++;
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

        //2024-08-15

        public static void WriteXLSHeader(Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet, System.Data.DataRow drMfr)
        {
            string OutputLine;

            WriteXLSLine(xlWorkSheet, "Purchase Order," + ",POTerms:," + drMfr["poterms"].ToString() + ",Buyer:," + drMfr["buyername"].ToString(), 1, true);
            WriteXLSLine(xlWorkSheet, "Issue Date," + drMfr["startdate"].ToString(), 2);
            WriteXLSLine(xlWorkSheet, "Ship Start Date," + drMfr["postartshipdate1"].ToString(), 3);
            WriteXLSLine(xlWorkSheet, "Ship Cancel Date," + drMfr["pocanceldate1"].ToString(), 4);
            WriteXLSLine(xlWorkSheet, "First Receipt,-", 5);
            WriteXLSLine(xlWorkSheet, "Last Receipt,-", 6);
            WriteXLSLine(xlWorkSheet, " ", 7);
            WriteXLSLine(xlWorkSheet, "AutoCancel Enabled," + drMfr["poautocancel"].ToString(), 8, true);

            OutputLine = "Orig. Line,Manufacturer,P.O.,Posted,Style,Description,Color,Color Code,size,UPC,Qty Ordered,Cost,Ext Cost,Receive,Color Override,Style Override,Closeout,SKU-Closeout,Backorder";
            WriteXLSLine(xlWorkSheet, OutputLine, 10, true);
        }

        //2024-08-15
        public static void WriteXLSFooter(Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet, System.Data.DataRow drMfr, System.Data.DataRow drCost, int SkuRow)
        {

            WriteXLSLine(xlWorkSheet, ",Lines,Pieces,Value", SkuRow + 1, true);
            WriteXLSLine(xlWorkSheet, "Original," + drCost["lines"].ToString() + "," + drCost["quantity"].ToString() + "," + drCost["cost"].ToString(), SkuRow + 2, true);
            WriteXLSLine(xlWorkSheet, "Current,0,0,0", SkuRow + 3, true);
            WriteXLSLine(xlWorkSheet, "Special Instructions," + drMfr["ponote"].ToString(), SkuRow + 4, true);
            WriteXLSLine(xlWorkSheet, " ", SkuRow + 5);

        }

        public static void WriteXLSLine(Microsoft.Office.Interop.Excel.Worksheet wksheet,HelperModels.POOutput POOutput, int Rowcount)
        {
            wksheet.Cells[Rowcount, 1] = POOutput.LineNumber;
            wksheet.Cells[Rowcount, 2] = POOutput.Manufacturer  ;
            wksheet.Cells[Rowcount, 3] = POOutput.PONumber  ;
            wksheet.Cells[Rowcount, 4] = POOutput.Posted  ;
            wksheet.Cells[Rowcount, 5] = POOutput.Style  ;
            wksheet.Cells[Rowcount, 6] = POOutput.Description  ;
            wksheet.Cells[Rowcount, 7] = POOutput.Color  ;
            wksheet.Cells[Rowcount, 8] = POOutput.ColorCode  ;
            wksheet.Cells[Rowcount, 9] = POOutput.Size  ;
            wksheet.Cells[Rowcount, 10] = POOutput.UPC  ;
            wksheet.Cells[Rowcount, 11] = POOutput.QtyOrdered  ;
            wksheet.Cells[Rowcount, 12] = POOutput.Cost  ;
            wksheet.Cells[Rowcount, 13] = POOutput.ExtCost  ;
            wksheet.Cells[Rowcount, 14] = POOutput.Receive  ;
            wksheet.Cells[Rowcount, 15] = POOutput.ColorOverride  ;
            wksheet.Cells[Rowcount, 16] = POOutput.StyleOverride  ;
            wksheet.Cells[Rowcount, 17] = POOutput.Closeout  ;
            wksheet.Cells[Rowcount, 18] = POOutput.SKUCloseout  ;
            wksheet.Cells[Rowcount, 19] = POOutput.Backorder;
        }

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
                Redded = CellAlertColumn.Split(Char.Parse(","));
                for (int xx=0; xx < Redded.Length-1; xx++)
                {          
                    //Console.WriteLine("redded [" + Redded[xx] + Rowcount.ToString() + ":" + Redded[xx] + Rowcount.ToString() + "]");
                    wksheet.Range[Redded[xx] + Rowcount.ToString() + ":" + Redded[xx] + Rowcount.ToString()].Font.Color = System.Drawing.Color.Red;
                }
            }
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
