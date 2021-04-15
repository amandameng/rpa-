using System;
using System.Collections.Generic;

using System.IO;
using System.Linq;
using System.Data;
using System.Text;

namespace ConsoleApp1.helpers
{
    class TxtParser
    {
        public static void FlexMNPo()
        {
            string a = "304.27 $";
            string b = a.Split('$', StringSplitOptions.RemoveEmptyEntries)[0].Trim();

            string file = @"D:\POC\伟创力\sample pdf\MN10009931.txt"; //with Revision
            //string file = @"D:\POC\伟创力\MN10009954.txt";


            string result = File.ReadAllText(file);

            string vendorKey = "VENDOR:";
            int vendorIndex = result.IndexOf(vendorKey);
            int poNumberIndex = result.IndexOf("PO Number:");
            int vendorEndIndex = vendorIndex + vendorKey.Length;

            // !! VENDOR
            /*
                FLEXTRONICS SALES
                UNIT 7D FINANCIAL LPARK LABUAN COMP JALAN MERDEKA, 87000 WP LABUAN, 87000 6087 - 453788
                MALAYSIA
            */
            string vendor = result.Substring(vendorEndIndex, poNumberIndex - vendorEndIndex).Trim();

            Console.WriteLine("========================VENDOR==============================");
            Console.WriteLine(vendor);

            string requestedByKey = "Requested By:";
            int resuqestedByIndex = result.IndexOf(requestedByKey);
            int ShipToIndex = result.IndexOf("SHIP TO:");
            int resuqestedByEndIndex = resuqestedByIndex + requestedByKey.Length;

            // poToRequestedBy 需要二次解析
            // PO Number: Order Date: Authorized By: Currency: Shipping Terms: Tax Id: Requested By:
            /*
                MN10010355 10/12/2020 Linda Redmond USD
                FCA - Free Carrier
                
                Lajeunesse, Paul
            */


            string poToRequestedBy = result.Substring(resuqestedByEndIndex, ShipToIndex - resuqestedByEndIndex).Trim();

            Console.WriteLine("========================poToRequestedBy==============================");
            Console.WriteLine(poToRequestedBy);
            //MN10009931 12/18/2020\r\n1 12/17/2020\r\nAngela Davis USD\r\nFCA - Free Carrier\r\n\r\nGu, Christine

            //MN10009954 10/28/2020 Angela Davis USD\r\nFCA - Free Carrier\r\n\r\nAmesquita, Alex


            string[] poToRequestedByArr = poToRequestedBy.Split('\n');
            int poToReqLength = poToRequestedByArr.Length;

            string RequestedBy = poToRequestedByArr[3].Trim();
            string TaxId = poToRequestedByArr[2].Trim();
            string ShippingTerms = poToRequestedByArr[1].Trim();

            Console.WriteLine("========================RequestedBy | TaxId | ShippingTerms==============================");
            Console.WriteLine(RequestedBy + "|" + TaxId + "|" + ShippingTerms);

            string poToCurrency = poToRequestedByArr[0].Trim();
            string[] poToCurrencyArr = poToCurrency.Split(' ');
            string PONumber = poToCurrencyArr[0];
            string OrderDate = poToCurrencyArr[1];
            string AuthorizedBy = string.Join(' ', poToCurrencyArr.Skip(2).Take(poToCurrencyArr.Length - 3));
            string Currency = poToCurrencyArr[poToCurrencyArr.Length - 1];

            Console.WriteLine("========================PONumber | OrderDate | AuthorizedBy | Currency==============================");
            Console.WriteLine(PONumber + "|" + OrderDate + "|" + AuthorizedBy + "|" + Currency);


            string BillToKey = "BILL TO:";
            string ShipToKey = "SHIP TO:";
            int ShipToEndIndex = ShipToIndex + ShipToKey.Length;
            int BillToIndex = result.IndexOf(BillToKey);
            int BillToEndIndex = BillToIndex + BillToKey.Length;

            string ShipTo = result.Substring(ShipToEndIndex, BillToIndex - ShipToEndIndex).Trim();

            int DescriptionNoIndex = result.IndexOf("Description	NO.");

            string BillTo = result.Substring(BillToEndIndex, DescriptionNoIndex - BillToEndIndex);

            Console.WriteLine("========================ShipTo | BillTo==============================");
            Console.WriteLine(ShipTo + "|" + BillTo);


        }

        public static void FlexMNPoV2()
        {
            string file = @"D:\云扩科技\RPA\流程市场\流程范例\project_workspace\项目6\MN10009931.txt";

            string result = File.ReadAllText(file);

            string PONumberKeyWord = "PO Number:";
            string CurrencyKeyword = "Currency:";
            string OrderDateKeyword = "Order Date:";
            string shippingTermsKeyword = "Shipping Terms:";

            string poNum = getEleValue(result, PONumberKeyWord);
            string currency = getEleValue(result, CurrencyKeyword);
            string OrderDate = getEleValue(result, OrderDateKeyword);
            string shippingTerms = getEleValue(result, shippingTermsKeyword);

            Console.WriteLine("{0}, {1}, {2}, {3}", poNum, currency, OrderDate, shippingTerms);
            Console.WriteLine("-----");

        }

        public static string getEleValue(string result, string PONumberKeyWord)
        {
            int PONumIndex = result.IndexOf(PONumberKeyWord);
            int PONumStrEndIndex = result.IndexOf("\r\n", PONumIndex);
            int POEndIndex = PONumIndex + PONumberKeyWord.Length;
            string PONum = result.Substring(POEndIndex, PONumStrEndIndex - POEndIndex).Trim();
            return PONum;
        }
            

        public static void FlexRMAPo()
        {
            //string file = @"D:\POC\伟创力\Purchase Order 4500594473.txt";
            string file = @"D:\POC\伟创力\sample pdf\Birgit-Purchase Order 4500602730.txt";

            string result = File.ReadAllText(file);
            string rmaKeyWord = "RMAMWTH-";
            int rmaKeyWordIndex = result.IndexOf(rmaKeyWord);

            // basicInfo keys: vendor, RequestedBy, TaxId, ShippingTerms, PONumber, OrderDate, AuthorizedBy, Currency, ShipTo, BillTo
            Dictionary<string, string> basicInfo = new Dictionary<string, string>();
            string rmaDescription = "";
            if (rmaKeyWordIndex != -1) // 需要处理的pdf
            {
                Console.WriteLine("-----RMA-----");
                int rmaKeyWordEndIndex = result.IndexOf("\r\n", rmaKeyWordIndex);
                rmaDescription = result.Substring(rmaKeyWordIndex, rmaKeyWordEndIndex - rmaKeyWordIndex).Trim();

                Console.WriteLine("-----rmaDescription----");
                Console.WriteLine(rmaDescription);
            }
            else
            {
                Console.WriteLine("-----Shippment----");
            }
            /*
             *  Your Vendor Number with us: 132466 FLEXTRONICS SALES AND MARKETING NORTH ASIA (L) LTD
                UNIT 7D FINANCIAL LPARK LABUAN COMP 87000 JALAN MERDEKA 87000 WP LABUAN MALAYSIA


                Please deliver to: Mitel Distribution
                2160 West Broadway, Suite #103 MESA AZ  85202
                UNITED STATES OF AMERICA

                Vendor Invoice Address: FLEXTRONICS SALES
                UNIT 7D FINANCIAL LPARK LABUAN COMP 87000 JALAN MERDEKA, 87000 WP LABUAN

             */
                string vendorKey = "Your Vendor Number with us:"; // Vendor
                int vendorIndex = result.IndexOf(vendorKey);
                int vendorEndIndex = vendorIndex + vendorKey.Length;

                string shipToKey = "Please deliver to:";  // shipTo
                int shipToIndex = result.IndexOf(shipToKey);
                int shipToEndIndex = shipToIndex + shipToKey.Length;

                //Vendor
                string vendor = result.Substring(vendorEndIndex, shipToIndex - vendorEndIndex).Trim();
                Console.WriteLine("--------Vendor----------");
                Console.WriteLine(vendor);

                string billToKey = "Vendor Invoice Address:";  //billTo
                int billToIndex = result.IndexOf(billToKey);
                int billToEndIndex = billToIndex + billToKey.Length;

                // ShipTo = Please deliver to:
                string ShipTo = result.Substring(shipToEndIndex, billToIndex - shipToEndIndex).Trim();
                Console.WriteLine("--------ShipTo----------");
                Console.WriteLine(ShipTo);


                string PONumberKey = "PO Number:";
                int PONumberIndex = result.IndexOf(PONumberKey);
                int PONumberEndIndex = result.IndexOf(PONumberKey) + PONumberKey.Length;
                int PONumberRowIndex = result.IndexOf("\r\n", PONumberIndex);

                string POAndCreatedOn = result.Substring(PONumberEndIndex, PONumberRowIndex - PONumberEndIndex);
                string[] POAndCreatedOnArr = POAndCreatedOn.Split(new string[] { "Created on:" }, StringSplitOptions.RemoveEmptyEntries);
                string PONumber = string.Empty;
                string createdOn = string.Empty;
                if (POAndCreatedOnArr.Length > 1)
                {
                    PONumber = POAndCreatedOnArr[0].Trim();
                    createdOn = POAndCreatedOnArr[1].Trim();
                }
                else
                {
                    PONumber = POAndCreatedOnArr[0].Trim();
                }
                Console.WriteLine("--------PONumber----------");
                Console.WriteLine(PONumber);
                Console.WriteLine("--------Created On----------");
                Console.WriteLine(createdOn);

                //billTo
                string billToRaw = result.Substring(billToEndIndex, PONumberIndex - billToEndIndex).Trim(); // 取Vendor Invoice Address: 和 PO Number:之间
                string billTo = billToRaw.Split(new string[] { "\n\n" }, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                Console.WriteLine("--------billTo----------");
                Console.WriteLine(billTo);


                /*
                 * Incoterms:  FCA SHIPPING POINT Currency:   USD
                   Payment Terms:  30 days from date of invoice
                 */
                string IncotermsKey = "Incoterms:";
                int IncotermsIndex = result.IndexOf(IncotermsKey);
                int IncotermsEndIndex = IncotermsIndex + IncotermsKey.Length;
                int IncotermsRowIndex = result.IndexOf("\r\n", IncotermsIndex);
                string IncotermsAndCurrency = result.Substring(IncotermsEndIndex, IncotermsRowIndex - IncotermsEndIndex); //"Incoterms:那一行
                string[] IncotermsCurrencyArr = IncotermsAndCurrency.Split(new string[] { "Currency:" }, StringSplitOptions.RemoveEmptyEntries);
                string Incoterms = string.Empty;
                string currency = string.Empty;
                if (IncotermsCurrencyArr.Length > 1)
                {
                    Incoterms = IncotermsCurrencyArr[0].Trim();
                    currency = IncotermsCurrencyArr[1].Trim();
                }
                else
                {
                    Incoterms = IncotermsCurrencyArr[0].Trim();
                }

                Console.WriteLine("--------Incoterms----------");
                Console.WriteLine(Incoterms);

                Console.WriteLine("--------Currency----------");
                Console.WriteLine(currency);

                string PaymentTermsKey = "Payment Terms:";
                int PaymentTermsIndex = result.IndexOf(PaymentTermsKey);
                int PaymentTermsEndIndex = PaymentTermsIndex + PaymentTermsKey.Length;

                int PaymentTermsRowIndex = result.IndexOf("\r\n", PaymentTermsIndex);


                string paymentTerms = result.Substring(PaymentTermsEndIndex, PaymentTermsRowIndex - PaymentTermsEndIndex).Trim(); //Payment Terms: 那一行查找
                Console.WriteLine("--------paymentTerms----------");
                Console.WriteLine(paymentTerms);


                string dateKey = "Date:";
                int dateIndex = result.IndexOf(dateKey);
                int DateEndIndex = dateIndex + dateKey.Length;

                int dateKeyNextLine = result.IndexOf("\r\n", dateIndex);
                string Date = result.Substring(DateEndIndex, dateKeyNextLine - DateEndIndex).Trim(); //Date那一行查找

                Console.WriteLine("--------Date----------");
                Console.WriteLine(Date);

                int orderItemsIndex;
                string orderTermsStart = "ORDER TERMS AND CONDITIONS";
                string frenchorderTermsStart = "CONDITIONS GÉNÉRALES D'ACHAT";
                string sepLine = "_______________________________________________________________________________________________________________________________";
                int sepLineIndex = result.IndexOf(sepLine, PaymentTermsIndex);

                int sepLineEndIndex = result.IndexOf("\r\n", sepLineIndex) + "\r\n".Length;
                
                orderItemsIndex = result.IndexOf(orderTermsStart, sepLineEndIndex); //sepLineEndIndex之后查找 orderTermsStart
                if (orderItemsIndex == -1)
                {
                    orderItemsIndex = result.IndexOf(frenchorderTermsStart, sepLineEndIndex); //sepLineEndIndex之后查找 frenchorderTermsStart
                }
                
            // basicInfo keys: vendor, RequestedBy, TaxId, ShippingTerms, PONumber, OrderDate, AuthorizedBy, Currency, ShipTo, BillTo
            basicInfo.Add("vendor", vendor);
                basicInfo.Add("RequestedBy", "");
                basicInfo.Add("TaxId", "");
                basicInfo.Add("ShippingTerms", Incoterms);
                basicInfo.Add("PONumber", PONumber);
                basicInfo.Add("OrderDate", Date);
                basicInfo.Add("AuthorizedBy", "");
                basicInfo.Add("Currency", currency);
                basicInfo.Add("ShipTo", ShipTo);
                basicInfo.Add("BillTo", billTo);


                // 后续的查询基于这一款字符串截取
                string itemsInfoStr = result.Substring(sepLineEndIndex, orderItemsIndex - sepLineEndIndex);

                string[] itemsInfoStrArr = itemsInfoStr.Split(new string[] { "\r\n", "_" }, StringSplitOptions.RemoveEmptyEntries);

                List<string> formatedStrList = new List<string> { };

                // ITEM\tMATERIAL/DES\tREQ DELVRY\tORD QTY\tBAL DUE\tUNIT\tPRICE/ UNIT\tNET VALUE
                // 00010\t50006478\t15.07.2020\t1\t1\tPiece\t31.83\t31.83 5340E IP PHONE
                foreach (string item in itemsInfoStrArr)
                {
                    string strippedItem = item.Trim();

                    if (strippedItem.Contains("\t"))
                    {
                        string shippingModeStr = "Shipping instructions:";
                        int shippingModeIndex = strippedItem.IndexOf(shippingModeStr);
                        if (shippingModeIndex != -1) {
                            int shippingModeEndIndex = shippingModeIndex + shippingModeStr.Length;
                            string shippingMode = strippedItem.Substring(shippingModeEndIndex, strippedItem.Length - shippingModeEndIndex).Trim();
                            basicInfo.Add("shippingMode", shippingMode);
                        basicInfo.ContainsKey("a");
                        }
                        string[] itemArr = strippedItem.Split(new string[] { "\t" }, StringSplitOptions.RemoveEmptyEntries);
                        if (!formatedStrList.Contains(strippedItem) && itemArr.Length == 8)
                        {
                            formatedStrList.Add(strippedItem);
                        }

                    }
                }

                // 搭建descriptions 数据表
                DataTable itemsDataTable = new DataTable();

                string[] columns = new string[] { "Description", "NO.", "Supplier Part ID", "Needed By", "Payment Terms", "Quantity", "Unit Price", "Subtotal" };
                foreach (string column in columns)
                {
                    itemsDataTable.Columns.Add(column, typeof(string));
                }

                int rowIndex = 0;
                foreach (string item in formatedStrList)
                {
                    string[] itemArr = item.Split(new string[] { "\t" }, StringSplitOptions.RemoveEmptyEntries);
                    if (rowIndex > 0)
                    {
                        // ITEM\tMATERIAL/DES\tREQ DELVRY\tORD QTY\tBAL DUE\tUNIT\tPRICE/ UNIT\tNET VALUE
                        // 00010\t50006478\t15.07.2020\t1\t1\tPiece\t31.83\t31.83 5340E IP PHONE

                        DataRow dr = itemsDataTable.NewRow();
                        string Description = itemArr[1];
                        string No = itemArr[0];
                        // string SupplierPart_ID = itemArr[1]
                        string NeededBy = itemArr[2];
                        string PaymentTerms = paymentTerms;
                        string Quantity = itemArr[3];
                        string UnitPrice = itemArr[6];
                        string subTotal = Convert.ToString(Convert.ToInt32(Quantity) * Convert.ToDouble(UnitPrice));

                        // get "Supplier Part ID"
                        string item7 = itemArr[7];
                        int item7Index = item7.IndexOf(" "); //31.83 5340E IP PHONE
                        string partOfSupplierID = item7.Substring(item7Index);
                        string supplierPartId = itemArr[1] + partOfSupplierID;

                        itemsDataTable.Rows.Add(new Object[] { Description, No, supplierPartId, NeededBy, PaymentTerms, Quantity, UnitPrice, subTotal });
                    }
                    rowIndex += 1;
                }
                Console.WriteLine("111");
            

        }

        public static void test()
        {
            try
            {
                //Dictionary<string, string>  basicInfo = new Dictionary<string, string>();
                string resultTxtFile = @"D:\POC\伟创力\Purchase Order 4500594473.txt";
                string result = File.ReadAllText(resultTxtFile);

                string rmaKeyWord = "RMAMWTH-";
                int rmaKeyWordIndex = result.IndexOf(rmaKeyWord);

                // basicInfo keys: vendor, RequestedBy, TaxId, ShippingTerms, PONumber, OrderDate, AuthorizedBy, Currency, ShipTo, BillTo
                Dictionary<string, string> basicInfo = new Dictionary<string, string>();

                if (rmaKeyWordIndex != -1) // 需要处理的pdf
                {

                    bool pdfNeedToParse = true;
                    /*
                     *  Your Vendor Number with us: 132466 FLEXTRONICS SALES AND MARKETING NORTH ASIA (L) LTD
                        UNIT 7D FINANCIAL LPARK LABUAN COMP 87000 JALAN MERDEKA 87000 WP LABUAN MALAYSIA


                        Please deliver to: Mitel Distribution
                        2160 West Broadway, Suite #103 MESA AZ  85202
                        UNITED STATES OF AMERICA

                        Vendor Invoice Address: FLEXTRONICS SALES
                        UNIT 7D FINANCIAL LPARK LABUAN COMP 87000 JALAN MERDEKA, 87000 WP LABUAN

                     */
                    string vendorKey = "Your Vendor Number with us:"; // Vendor
                    int vendorIndex = result.IndexOf(vendorKey);
                    int vendorEndIndex = vendorIndex + vendorKey.Length;

                    string shipToKey = "Please deliver to:";  // shipTo
                    int shipToIndex = result.IndexOf(shipToKey);
                    int shipToEndIndex = shipToIndex + shipToKey.Length;

                    //Vendor
                    string vendor = result.Substring(vendorEndIndex, shipToIndex - vendorEndIndex).Trim();
                    Console.WriteLine("--------Vendor----------");
                    Console.WriteLine(vendor);

                    string billToKey = "Vendor Invoice Address:";  //billTo
                    int billToIndex = result.IndexOf(billToKey);
                    int billToEndIndex = billToIndex + billToKey.Length;

                    // ShipTo = Please deliver to:
                    string ShipTo = result.Substring(shipToEndIndex, billToIndex - shipToEndIndex).Trim();
                    Console.WriteLine("--------ShipTo----------");
                    Console.WriteLine(ShipTo);


                    string PONumberKey = "PO Number:";
                    int PONumberIndex = result.IndexOf(PONumberKey);
                    int PONumberEndIndex = result.IndexOf(PONumberKey) + PONumberKey.Length;
                    int PONumberRowIndex = result.IndexOf("\r\n", PONumberIndex);

                    string POAndCreatedOn = result.Substring(PONumberEndIndex, PONumberRowIndex - PONumberEndIndex);
                    string[] POAndCreatedOnArr = POAndCreatedOn.Split(new string[] { "Created on:" }, StringSplitOptions.RemoveEmptyEntries);
                    string PONumber = string.Empty;
                    string createdOn = string.Empty;
                    if (POAndCreatedOnArr.Length > 1)
                    {
                        PONumber = POAndCreatedOnArr[0].Trim();
                        createdOn = POAndCreatedOnArr[1].Trim();
                    }
                    else
                    {
                        PONumber = POAndCreatedOnArr[0].Trim();
                    }
                    Console.WriteLine("--------PONumber----------");
                    Console.WriteLine(PONumber);
                    Console.WriteLine("--------Created On----------");
                    Console.WriteLine(createdOn);

                    //billTo
                    string billToRaw = result.Substring(billToEndIndex, PONumberIndex - billToEndIndex).Trim(); // 取Vendor Invoice Address: 和 PO Number:之间
                    string billTo = billToRaw.Split(new string[] { "\n\n" }, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    Console.WriteLine("--------billTo----------");
                    Console.WriteLine(billTo);


                    /*
                     * Incoterms:  FCA SHIPPING POINT Currency:   USD
                       Payment Terms:  30 days from date of invoice
                     */
                    string IncotermsKey = "Incoterms:";
                    int IncotermsIndex = result.IndexOf(IncotermsKey);
                    int IncotermsEndIndex = IncotermsIndex + IncotermsKey.Length;
                    int IncotermsRowIndex = result.IndexOf("\r\n", IncotermsIndex);
                    string IncotermsAndCurrency = result.Substring(IncotermsEndIndex, IncotermsRowIndex - IncotermsEndIndex); //"Incoterms:那一行
                    string[] IncotermsCurrencyArr = IncotermsAndCurrency.Split(new string[] { "Currency:" }, StringSplitOptions.RemoveEmptyEntries);
                    string Incoterms = string.Empty;
                    string currency = string.Empty;
                    if (IncotermsCurrencyArr.Length > 1)
                    {
                        Incoterms = IncotermsCurrencyArr[0].Trim();
                        currency = IncotermsCurrencyArr[1].Trim();
                    }
                    else
                    {
                        Incoterms = IncotermsCurrencyArr[0].Trim();
                    }

                    Console.WriteLine("--------Incoterms----------");
                    Console.WriteLine(Incoterms);

                    Console.WriteLine("--------Currency----------");
                    Console.WriteLine(currency);



                    string PaymentTermsKey = "Payment Terms:";
                    int PaymentTermsIndex = result.IndexOf(PaymentTermsKey);
                    int PaymentTermsEndIndex = PaymentTermsIndex + PaymentTermsKey.Length;

                    int PaymentTermsRowIndex = result.IndexOf("\r\n", PaymentTermsIndex);


                    string paymentTerms = result.Substring(PaymentTermsEndIndex, PaymentTermsRowIndex - PaymentTermsEndIndex).Trim(); //Payment Terms: 那一行查找
                    Console.WriteLine("--------paymentTerms----------");
                    Console.WriteLine(paymentTerms);


                    string dateKey = "Date:";
                    int dateIndex = result.IndexOf(dateKey);
                    int DateEndIndex = dateIndex + dateKey.Length;

                    int dateKeyNextLine = result.IndexOf("\r\n", dateIndex);
                    string Date = result.Substring(DateEndIndex, dateKeyNextLine - DateEndIndex).Trim(); //Date那一行查找

                    Console.WriteLine("--------Date----------");
                    Console.WriteLine(Date);

                    int rmaKeyWordNextLine = result.IndexOf("\r\n", rmaKeyWordIndex);
                    string RMAKeyWordRow = result.Substring(rmaKeyWordIndex, rmaKeyWordNextLine - rmaKeyWordIndex);

                    Console.WriteLine("--------RMAKeyWordDescp----------");
                    Console.WriteLine(RMAKeyWordRow);


                    string orderTermsStart = "ORDER TERMS AND CONDITIONS";
                    int orderItemsIndex = result.IndexOf(orderTermsStart, rmaKeyWordIndex); //rmaKeyWordIndex之后查找


                    // basicInfo keys: vendor, RequestedBy, TaxId, ShippingTerms, PONumber, OrderDate, AuthorizedBy, Currency, ShipTo, BillTo
                    basicInfo.Add("vendor", vendor);
                    basicInfo.Add("RequestedBy", "");
                    basicInfo.Add("TaxId", "");
                    basicInfo.Add("ShippingTerms", Incoterms);
                    basicInfo.Add("PONumber", PONumber);
                    basicInfo.Add("OrderDate", Date);
                    basicInfo.Add("AuthorizedBy", "");
                    basicInfo.Add("Currency", currency);
                    basicInfo.Add("ShipTo", ShipTo);
                    basicInfo.Add("BillTo", billTo);


                    // 后续的查询基于这一款字符串截取
                    string itemsInfoStr = result.Substring(rmaKeyWordIndex, orderItemsIndex - rmaKeyWordIndex);

                    string[] itemsInfoStrArr = itemsInfoStr.Split(new string[] { "\r\n", "_" }, StringSplitOptions.RemoveEmptyEntries);

                    List<string> formatedStrList = new List<string> { };

                    // ITEM\tMATERIAL/DES\tREQ DELVRY\tORD QTY\tBAL DUE\tUNIT\tPRICE/ UNIT\tNET VALUE
                    // 00010\t50006478\t15.07.2020\t1\t1\tPiece\t31.83\t31.83 5340E IP PHONE
                    foreach (string item in itemsInfoStrArr)
                    {
                        string strippedItem = item.Trim();

                        if (strippedItem.Contains("\t"))
                        {
                            string[] itemArr = strippedItem.Split(new string[] { "\t" }, StringSplitOptions.RemoveEmptyEntries);
                            if (!formatedStrList.Contains(strippedItem) && itemArr.Length == 8)
                            {
                                formatedStrList.Add(strippedItem);
                            }

                        }
                    }

                    // 搭建descriptions 数据表
                    DataTable itemsDataTable = new DataTable();

                    string[] columns = new string[] { "Description", "NO.", "Supplier Part ID", "Needed By", "Payment Terms", "Quantity", "Unit Price", "Subtotal" };
                    foreach (string column in columns)
                    {
                        itemsDataTable.Columns.Add(column, typeof(string));
                    }

                    int rowIndex = 0;
                    foreach (string item in formatedStrList)
                    {
                        string[] itemArr = item.Split(new string[] { "\t" }, StringSplitOptions.RemoveEmptyEntries);
                        if (rowIndex > 0)
                        {
                            // ITEM\tMATERIAL/DES\tREQ DELVRY\tORD QTY\tBAL DUE\tUNIT\tPRICE/ UNIT\tNET VALUE
                            // 00010\t50006478\t15.07.2020\t1\t1\tPiece\t31.83\t31.83 5340E IP PHONE

                            DataRow dr = itemsDataTable.NewRow();
                            string Description = itemArr[1];
                            string No = itemArr[0];
                            // string SupplierPart_ID = itemArr[1]
                            string NeededBy = itemArr[2];
                            string PaymentTerms = paymentTerms;
                            string Quantity = itemArr[3];
                            string UnitPrice = itemArr[6];
                            string subTotal = Convert.ToString(Convert.ToInt32(Quantity) * Convert.ToDouble(UnitPrice));

                            // get "Supplier Part ID"
                            string item7 = itemArr[7];
                            int item7Index = item7.IndexOf(" "); //31.83 5340E IP PHONE
                            string partOfSupplierID = item7.Substring(item7Index);
                            string supplierPartId = itemArr[1] + partOfSupplierID;

                            itemsDataTable.Rows.Add(new Object[] { Description, No, supplierPartId, NeededBy, PaymentTerms, Quantity, UnitPrice, subTotal });
                        }
                        rowIndex += 1;
                    }
                    Console.WriteLine("111");
                }
                else
                {
                    bool pdfNeedToParse = false;
                    Console.WriteLine("-----不是需要处理的pdf");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw;
            }

        }

        public static void mciPdf()
        {
            string file = @"D:\chrome downloads\result.txt";

            string result = File.ReadAllText(file);
            string materialKeyWord = "Material";
            int mtrKeyWordIndex = result.IndexOf(materialKeyWord);
            Console.WriteLine("-----Material-----");
            int mtrKeyWordEndIndex = result.IndexOf("\r\n", mtrKeyWordIndex);
            
            string materialLine = result.Substring(mtrKeyWordIndex, mtrKeyWordEndIndex - mtrKeyWordIndex);
            string material = materialLine.Split(" ", StringSplitOptions.RemoveEmptyEntries)[1];
            Console.WriteLine(material);

            string delDateStr = "Del.date";
            int delDateIndex = result.IndexOf(delDateStr);
            if (delDateIndex == -1) {
                delDateStr = "Ship.date";
                delDateIndex = result.IndexOf(delDateStr);
            }

            string schedulesText = "Schedules are usually sent weekly";
            int schedulesIndex = result.LastIndexOf(schedulesText);
            
            string materialItems = result.Substring(delDateIndex, schedulesIndex - delDateIndex);
            Console.WriteLine(materialItems);

            DataTable itemsDataTable = new DataTable();

            string[] columns = new string[] { "DelDate", "Qty", "Change", "Agreed cum" };
            foreach (string column in columns)
            {
                itemsDataTable.Columns.Add(column, typeof(string));
            }

            string[] eachRowArr = materialItems.Split("\r\n");
            
            foreach(string item in eachRowArr)
            {
                if (!string.IsNullOrEmpty(item.Trim()))
                {
                    string[] itemArr = item.Split(" ", StringSplitOptions.RemoveEmptyEntries);
                    DateTime outDate;
                    bool isDate = DateTime.TryParse(itemArr[0], out outDate);
                    if(itemArr[0] != delDateStr && itemArr.Length == columns.Length && isDate)
                    {
                        itemsDataTable.Rows.Add(itemArr);
                    }
                }
            }

            foreach(DataRow dr in itemsDataTable.Rows)
            {
                Console.WriteLine("{0}, {1}, {2}, {3}", dr[0], dr[1], dr[2], dr[3]);
            }



        }


        //字符转ASCII码
        public static int Asc(string character)
        {
            if (character.Length == 1)
            {
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                int intAsciiCode = (int)asciiEncoding.GetBytes(character)[0];
                return (intAsciiCode);
            }
            else
            {
                throw new Exception("Character is not valid.");
            }

        }

        //ASCII码转字符：

        public static string Chr(int asciiCode)
        {
            if (asciiCode >= 0 && asciiCode <= 255)
            {
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                byte[] byteArray = new byte[] { (byte)asciiCode };
                string strCharacter = asciiEncoding.GetString(byteArray);
                return (strCharacter);
            }
            else
            {
                throw new Exception("ASCII Code is not valid.");
            }
        }
        

        public static void getEmail()
        {
            string origStr = @"
	1:一般电话			13502116233张钰健					
	2:移动电话			022-28668922					
	4:电子邮件			hlys2011@163.com					
									
									
									
									
									
									
									
									
									
									
";
            string[] lineArr = origStr.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string line in lineArr)
            {
                if (!string.IsNullOrEmpty(line.Trim()))
                {
                    Console.WriteLine(line);
                    string[] tabArr = line.Split(new string[] { "\t" }, StringSplitOptions.None);
                    if (tabArr[1].Contains("电子邮件")){
                        Console.WriteLine(tabArr[4]);
                    }
                }
               
            }

        }
    }
}


