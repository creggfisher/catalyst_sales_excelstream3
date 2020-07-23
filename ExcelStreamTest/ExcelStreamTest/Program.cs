using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper;
using NPOI.SS.UserModel;

namespace ExcelStreamTest
{
    class Program
    {

        void PrintToEbookFlavor1(List<SalesDetailInput> records)
        {
            // Do stuff here
        }
        void PrintToEbookFlavor2(List<SalesDetailInput> records)
        {
            // Do stuff here
        }


        // xlsx, good
        static List<SalesDetailInput> Transform3M(string filePath)
        {
            Console.WriteLine("Transform3M initiated " + DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt"));
            var custlookup = new Dictionary<string,int>();
            custlookup.Add("AU", 109707);
            custlookup.Add("CA", 106532);
            custlookup.Add("US", 100761);
            CultureInfo enUS = new CultureInfo("en-US");
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            using (FileStream fs = File.OpenRead(filePath))
            {
                IWorkbook wb = WorkbookFactory.Create(fs);
                var sheet = wb.GetSheetAt(0);
                for (int i = 1; i <= sheet.LastRowNum; i++)
                {
                    var r = sheet.GetRow(i);
                    var m = r.GetCell(0).DateCellValue;
                    var mpp = m.AddMonths(1);
                    var qty = (int)r.GetCell(10).NumericCellValue;
                    if (qty == 0) { continue; }
                    var total = (decimal)r.GetCell(11).NumericCellValue;
                    var isSale = r.GetCell(10).NumericCellValue > 0;
                    int custid;
                    if (!custlookup.TryGetValue(r.GetCell(5).StringCellValue.Trim(), out custid)) custlookup.TryGetValue("US", out custid);
                    records.Add(new SalesDetailInput
                    {
                        BilltoCustomerId = custid,
                        ShipToCustomerId = custid,
                        PostPeriod = (mpp.Year * 100) + mpp.Month,
                        SopTypeId = isSale ? 634 : 635,
                        PONumber = m.ToString("MMMM yyyy", DateTimeFormatInfo.InvariantInfo),
                        FiscalYear = m.Year,
                        FiscalMonth = m.Month,
                        ListPrice = Math.Abs((decimal)(r.GetCell(8).NumericCellValue / qty)),
                        SalesPrice = Math.Abs((decimal)(total / qty)),
                        Ean = r.GetCell(2).StringCellValue.Trim(),
                        GrossSalesUnits = isSale ? qty : 0,
                        GrossRtnUnits = isSale ? 0 : qty,
                        GrossSalesDols = isSale ? total : 0,
                        GrossRtnDols = isSale ? 0 : -1 * Math.Abs(total),
                        DiscRate = Math.Abs((decimal)(r.GetCell(9).NumericCellValue / 100)),
                        SourceRowNumber = i + 1,
                    });
                }
            }
            Console.WriteLine("Transform3M completed " + DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt"));
            return records;
        }

        // xlsx
        static List<SalesDetailInput> TransformBakerTaylor(string filePath)
        {
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            // stuff here
            return records;
        }

        // csv, good
        static List<SalesDetailInput> TransformBarnesNoble(string filePath)
        {
            Console.WriteLine("TransformBarnesNoble initiated " + DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt"));
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            // stuff here
            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                csv.Read();
                csv.ReadHeader();
                var i = 0;
                while (csv.Read())
                {
                    i++;
                    if (i == 1) { continue;  }
                    var Ean = csv.GetField(1);
                    if (Ean == null || Ean.Trim() == "") { continue; }
                    var m = csv.GetField<DateTime>(27);
                    var mpp = m.AddMonths(1);
                    var qty = csv.GetField<int>(19);
                    var isSale = qty > 0 ? true : false;
                    records.Add(new SalesDetailInput
                    {
                        BilltoCustomerId = 94024,
                        ShipToCustomerId = 94024,
                        PostPeriod = (mpp.Year * 100) + mpp.Month,
                        SopTypeId = isSale ? 634 : 635,
                        FiscalYear = m.Year,
                        FiscalMonth = m.Month,
                        InvoiceDate = m,
                        PONumber = m.ToString("MMMM yyyy", DateTimeFormatInfo.InvariantInfo),
                        Ean = csv.GetField(1),
                        OldSSDocnumber = csv.GetField(10),
                        GrossSalesUnits = qty,
                        GrossRtnUnits = csv.GetField<int>(20),
                        USListPrice = csv.GetField<decimal>(18),
                        ListPrice = csv.GetField<decimal>(15),
                        SalesPrice = csv.GetField<decimal>(22),
                        GrossSalesDols = isSale ? csv.GetField<decimal>(25) : 0,
                        GrossRtnDols = isSale ? 0 : csv.GetField<decimal>(18),
                        SourceRowNumber = i,
                    });
                }
            }

            Console.WriteLine("TransformBarnesNoble completed " + DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt"));
            return records;
        }

        // xls, good
        static List<SalesDetailInput> TransformBolinda(string filePath)
        {
            Console.WriteLine("TransformBolinda initiated " + DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt"));
            CultureInfo enUS = new CultureInfo("en-US");
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            using (FileStream fs = File.OpenRead(filePath))
            {
                IWorkbook wb = WorkbookFactory.Create(fs);
                var sheet = wb.GetSheetAt(0);
                var month = sheet.GetRow(0).GetCell(0).StringCellValue.Split(' ').Last();
                var year = sheet.GetRow(1).GetCell(0).StringCellValue.Split('/').Last();
                var m = DateTime.ParseExact(month + " " + year, "MMMM yyyy", enUS, DateTimeStyles.None);
                var postPeriod = (m.Year * 100) + m.Month;
                if (postPeriod == 201906) {
                    m = m.AddMonths(4);
                    postPeriod = (m.Year * 100) + m.Month;
                } else {
                    Console.WriteLine("post period identified: " + postPeriod);
                }
                var poNumber = m.ToString("MMMM yyyy", DateTimeFormatInfo.InvariantInfo);
                var fiscalYear = m.Year;
                var fiscalMonth = m.Month;
                for (int i = 6; i <= sheet.LastRowNum; i++)
                {
                    var r = sheet.GetRow(i);
                    string ean;
                    try {
                        ean = r.GetCell(1).StringCellValue;
                    } catch (Exception e) {
                        continue;
                    }
                    if (ean == null || ean == "") { continue; }
                    var qty = (int)r.GetCell(9).NumericCellValue;
                    if (qty == 0) { continue; }
                    var total = (decimal)r.GetCell(10).NumericCellValue;
                    var isSale = qty > 0;
                    records.Add(new SalesDetailInput
                    {
                        BilltoCustomerId = 115248,
                        ShipToCustomerId = 115248,
                        PostPeriod = postPeriod,
                        SopTypeId = isSale ? 634 : 635,
                        PONumber = poNumber,
                        FiscalYear = fiscalYear,
                        FiscalMonth = fiscalMonth,
                        USListPrice = Math.Abs((decimal)r.GetCell(7).NumericCellValue),
                        ListPrice = Math.Abs((decimal)r.GetCell(7).NumericCellValue),
                        SalesPrice = Math.Abs((decimal)(r.GetCell(8).NumericCellValue / qty)),
                        Ean = ean,
                        GrossSalesUnits = isSale ? qty : 0,
                        GrossRtnUnits = isSale ? 0 : qty,
                        GrossSalesDols = isSale ? total : 0,
                        GrossRtnDols = isSale ? 0 : total,
                        DiscRate = Math.Abs((decimal)(r.GetCell(8).NumericCellValue / 100)),
                        SourceRowNumber = i + 1,
                    });
                }
            }
            Console.WriteLine("TransformBolinda completed " + DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt"));
            return records;
        }

        // xlsx
        static List<SalesDetailInput> TransformComixology(string filePath)
        {
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            // stuff here
            return records;
        }

        // xlsx
        static List<SalesDetailInput> TransformCopia(string filePath)
        {
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            // stuff here
            return records;
        }

        // xlsx
        static List<SalesDetailInput> TransformCreateSpacePOD(string filePath)
        {
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            // stuff here
            return records;
        }

        // csv
        static List<SalesDetailInput> TransformDriveThruRPG(string filePath)
        {
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            // stuff here
            return records;
        }

        // unknown
        static List<SalesDetailInput> TransformEpic(string filePath)
        {
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            // stuff here
            return records;
        }

        // xls
        static List<SalesDetailInput> TransformFollett(string filePath)
        {
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            // stuff here
            return records;
        }

        // unknown
        static List<SalesDetailInput> TransformFutian(string filePath)
        {
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            // stuff here
            return records;
        }

        // csv
        static List<SalesDetailInput> TransformGoogleEarning(string filePath)
        {
            Console.WriteLine("TransformGoogleEarning initiated " + DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt"));
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            var custlookup = new Dictionary<string,int>();
            custlookup.Add("US", 94025);
            custlookup.Add("AU", 98701);
            custlookup.Add("NZ", 98701);
            custlookup.Add("CA", 98700);
            custlookup.Add("GB", 98702);
            custlookup.Add("BE", 99724);
            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                csv.Configuration.Delimiter = "\t";
                csv.Read();
                csv.ReadHeader();
                var i = 1;
                while (csv.Read())
                {
                    i++;
                    var country = csv.GetField(16);
                    if (country == "") { continue; }
                    int custid;
                    if (!custlookup.TryGetValue(country, out custid)) custlookup.TryGetValue("BE", out custid);
                    var m = csv.GetField<DateTime>(0);
                    var mpp = m.AddMonths(1);
                    var qty = csv.GetField<int>(6);
                    var total = csv.GetField<decimal>(20);
                    if (qty == 0) { continue; }
                    var isSale = qty > 0 ? true : false;
                    var listPrice = csv.GetField<decimal>(12);
                    var currConvRate = csv.GetField(21) == "" ? 0 : csv.GetField<decimal>(21);
                    records.Add(new SalesDetailInput
                    {
                        BilltoCustomerId = custid,
                        ShipToCustomerId = custid,
                        PostPeriod = (mpp.Year * 100) + mpp.Month,
                        SopTypeId = isSale ? 634 : 635,
                        FiscalYear = m.Year,
                        FiscalMonth = m.Month,
                        InvoiceDate = csv.GetField<DateTime>(1),
                        PONumber = m.ToString("MMMM yyyy", DateTimeFormatInfo.InvariantInfo),
                        Ean = csv.GetField(7),
                        OldSSDocnumber = csv.GetField(2),
                        GrossSalesUnits = isSale ? qty : 0,
                        GrossRtnUnits = isSale ? 0 : qty,
                        ListPrice = listPrice,
                        SalesPrice = Math.Abs(total),
                        GrossSalesDols = isSale ? total : 0,
                        GrossRtnDols = isSale ? 0 : total,
                        CurrConvRate = currConvRate,
                        SourceRowNumber = i,
                    });
                }
            }
            // PrintToEbookFlavor1(records);
            Console.WriteLine("TransformGoogleEarning completed " + DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss.fff tt"));
            return records;
        }

        // unknown
        static List<SalesDetailInput> TransformHummingbird(string filePath)
        {
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            // stuff here
            return records;
        }

        // xls
        static List<SalesDetailInput> TransformIngramBasic(string filePath)
        {
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            // stuff here
            return records;
        }

        // xls, use TransformIngramBasic
        static List<SalesDetailInput> TransformIngramAU(string filePath, int customerId)
        {
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            // stuff here
            return records;
        }

        // xls, use TransformIngramBasic
        static List<SalesDetailInput> TransformIngramGBP(string filePath)
        {
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            // stuff here
            return records;
        }

        // xls, use TransformIngramBasic
        static List<SalesDetailInput> TransformIngramUS(string filePath)
        {
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            // stuff here
            return records;
        }

        // xlsx
        static List<SalesDetailInput> TransformInkling(string filePath)
        {
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            // stuff here
            return records;
        }

        // xlsx
        static List<SalesDetailInput> TransformKobo(string filePath)
        {
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            // stuff here
            return records;
        }

        // xlsx
        static List<SalesDetailInput> TransformOverdrive(string filePath)
        {
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            // stuff here
            return records;
        }

        // xlsx
        static List<SalesDetailInput> TransformSSIndia(string filePath)
        {
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            // Get the appropriate sheet name to pull from
            var sheetName = "AMP source file";
            // var excel = new ExcelQueryFactory(filePath);
            // var sheets = excel.GetWorksheetNames();
            // if (sheets.Count() == 1)
            // {
            //     sheetName = sheets.First();
            // }
            // // Load data from spreadsheet
            // var rows = from c in excel.WorksheetRangeNoHeader("A2", "K" + (1+excel.Worksheet(sheetName).Count()), sheetName) select c;
            // var count = 0;
            // decimal conversionRate = 1;
            // var sourceRowNumber = 1;
            // // Read in initial data
            // foreach (var row in rows)
            // {
            //     sourceRowNumber++;
            //     // Special handling for rows that have an empty first column
            //     if (row[0] == null || row[0] == "") {
            //         // Gather conversion rate if it is the correct row
            //         if (row[7].ToString().Trim().ToLower() == "gbp") {
            //             conversionRate = row[10].Cast<decimal>();
            //         }
            //         continue;
            //     }
            //     // Parse row into preliminary set of data
            //     count++;
            //     var qty = Convert.ToInt16(row[9].ToString().Replace(",", ""));
            //     var year = row[0].Cast<int>();
            //     var month = row[1].Cast<int>();
            //     var rec = new SalesDetailInput(){
            //         FiscalYear = year,
            //         FiscalMonth = month,
            //         PostPeriod = (year * 100) + month,
            //         Ean = row[2].ToString().Trim(),
            //         BilltoCustomerId = 105912, // Magic number
            //         ShipToCustomerId = 105912, // Magic number
            //         SourceRowNumber = sourceRowNumber,
            //         GrossSalesUnits = qty > 0 ? 0 : -1 * qty,
            //         GrossRtnUnits = qty < 0 ? 0 : -1 * qty,
            //         GrossNativeCurrencyAmount = -1 * Convert.ToDecimal(row[8].ToString().Replace(",", "")),
            //     };
            //     records.Add(rec);
            // }
            // // Now that conversion rate has been identified we can perform a second pass
            // // This round allows us to create USD dollar values
            // foreach (var rec in records)
            // {
            //     var qty = rec.GrossSalesUnits > 0 ? rec.GrossSalesUnits : rec.GrossRtnUnits;
            //     // Exit early scenario
            //     if (qty == 0) { continue;  }
            //     // Compute USD values
            //     var usdAmount = rec.GrossNativeCurrencyAmount / conversionRate;
            //     var unitUSD = Math.Abs(usdAmount ?? 0) / Math.Abs(qty ?? 1);
            //     rec.SalesPrice = unitUSD;
            //     rec.CurrConvRate = conversionRate;
            //     if (rec.GrossSalesUnits > 0) {
            //         rec.GrossSalesDols = usdAmount;
            //     } else {
            //         rec.GrossRtnDols = usdAmount;
            //     }
            //     if (rec.SourceRowNumber <= 11) {
            //         Console.WriteLine("Test SS India Ean: {0}, GrossUSDAmount: {1:C2}", rec.Ean, usdAmount);
            //     }
            // }
            // Console.WriteLine("SS India rows: " + count.ToString() + "/" + records.Count);
            return records;
        }

        // unknown
        static List<SalesDetailInput> TransformSSUK(string filePath)
        {
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            // stuff here
            return records;
        }

        // unknown
        static List<SalesDetailInput> TransformSSUS(string filePath)
        {
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            // stuff here
            return records;
        }

        // unknown
        static List<SalesDetailInput> TransformSSRupi(string filePath)
        {
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            // stuff here
            return records;
        }

        // xlsx
        static List<SalesDetailInput> TransformVearsa(string filePath)
        {
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            // stuff here
            return records;
        }

        // xlsx
        static List<SalesDetailInput> TransformWheelers(string filePath)
        {
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            // stuff here
            return records;
        }

        // unknown
        static List<SalesDetailInput> TransformYuzu(string filePath)
        {
            List<SalesDetailInput> records = new List<SalesDetailInput>();
            // stuff here
            return records;
        }

        static void Main(string[] args)
        {
            Console.WriteLine("ExcelStreamTest initiated " + DateTime.Now);
            List<SalesDetailInput> testrecords = new List<SalesDetailInput>();
            var testcount = 0;

            // 3M --------------------------------------------------------------
            testrecords = Transform3M("C:\\Users\\JSullard\\projects\\justin-misc\\AMUTECH-669\\files\\from_sales_team\\eBook Samples\\3M.xlsx");
            foreach (var r in testrecords)
            {
                testcount++;
                if (testcount <= 10)
                {
                    Console.WriteLine("  row[" + testcount + "]: " + r.PostPeriod + ", " + r.Ean + ", " + r.GrossSalesUnits + ", " + r.PONumber);
                }
            }
            Console.WriteLine("test 3M records: " + testcount.ToString() + "/" + testrecords.Count);
            testrecords.Clear();
            testcount = 0;

            // Baker & Taylor --------------------------------------------------

            // Barnes & Noble --------------------------------------------------
            testrecords = TransformBarnesNoble("C:\\Users\\JSullard\\projects\\justin-misc\\AMUTECH-669\\files\\from_sales_team\\eBook Samples\\BN.csv");
            foreach (var r in testrecords)
            {
                testcount++;
                if (testcount <= 10)
                {
                    Console.WriteLine("  row[" + testcount + "]: " + r.PostPeriod + ", " + r.Ean + ", " + r.GrossSalesUnits + ", " + r.PONumber + ", " + r.SourceRowNumber);
                }
            }
            Console.WriteLine("test Barnes & Noble records: " + testcount.ToString() + "/" + testrecords.Count);
            testrecords.Clear();
            testcount = 0;

            // Bolinda ---------------------------------------------------------
            testrecords = TransformBolinda("C:\\Users\\JSullard\\projects\\justin-misc\\AMUTECH-669\\files\\from_sales_team\\eBook Samples\\BolindaDigital_Andrews Mcmeel Publishing US_201906_QTR_Sales.xls");
            foreach (var r in testrecords)
            {
                testcount++;
                if (testcount <= 10)
                {
                    Console.WriteLine("  row[" + testcount + "]: " + r.PostPeriod + ", " + r.Ean + ", " + r.GrossSalesUnits + ", " + r.PONumber);
                }
            }
            Console.WriteLine("test Bolinda records: " + testcount.ToString() + "/" + testrecords.Count);
            testrecords.Clear();
            testcount = 0;

            // Comixology ------------------------------------------------------
            // Copia -----------------------------------------------------------
            // CreateSpacePOD --------------------------------------------------
            // DriveThruRPG ----------------------------------------------------
            // Epic ------------------------------------------------------------
            // Follett ---------------------------------------------------------
            // Futian ----------------------------------------------------------

            // GoogleEarningsReport --------------------------------------------
            testrecords = TransformGoogleEarning("C:\\Users\\JSullard\\projects\\justin-misc\\AMUTECH-669\\files\\from_sales_team\\eBook Samples\\GoogleEarningsReport (38).csv");
            foreach (var r in testrecords)
            {
                testcount++;
                if (testcount <= 10)
                {
                    Console.WriteLine("  row[" + testcount + "]: " + r.PostPeriod + ", " + r.Ean + ", " + r.GrossSalesUnits + ", " + r.PONumber + ", " + r.SourceRowNumber);
                }
            }
            Console.WriteLine("test Google records: " + testcount.ToString() + "/" + testrecords.Count);
            testrecords.Clear();
            testcount = 0;

            // Hummingbird -----------------------------------------------------
            // IngramAU --------------------------------------------------------
            // IngramGBP -------------------------------------------------------
            // IngramUS --------------------------------------------------------
            // Inkling ---------------------------------------------------------
            // Iverse ----------------------------------------------------------
            // Kobo ------------------------------------------------------------
            // Overdrive -------------------------------------------------------
            // SSIndia ---------------------------------------------------------
            // SSUK ------------------------------------------------------------
            // SSUS ------------------------------------------------------------
            // SSRupi ----------------------------------------------------------
            // Vearsa ----------------------------------------------------------
            // Wheelers --------------------------------------------------------
            // Yuzu ------------------------------------------------------------

            // -----------------------------------------------------------------
            Console.Write("ExcelStreamTest completed (" + DateTime.Now + "), press any key to exit");
            Console.ReadKey();
        }

        public class SalesDetailInput
        {
            public Nullable<int> BilltoCustomerId { get; set; }
            public Nullable<int> CountryId { get; set; }
            public Nullable<int> CreditReturnReasonId { get; set; }
            public Nullable<int> CurrencyTypeId { get; set; }
            public Nullable<int> CustomerSalesRepId { get; set; }
            public Nullable<int> MarketSegementId { get; set; }
            public Nullable<int> OrderTypeId { get; set; }
            public Nullable<int> PostPeriod { get; set; }
            public Nullable<int> ProductId { get; set; }
            public Nullable<int> ProductClassId { get; set; }
            public Nullable<int> ProductSubClassId { get; set; }
            public Nullable<int> RoyaltyChannelId { get; set; }
            public Nullable<int> SalesLineItemTypeId { get; set; }
            public Nullable<int> SetProductId { get; set; }
            public Nullable<int> ShipToCustomerId { get; set; }
            public Nullable<int> ShiptoStateId { get; set; }
            public Nullable<int> SopTypeId { get; set; }
            public Nullable<int> WarehouseId { get; set; }
            public string BillToCustomerCispubId { get; set; }
            public string ShipToCustomerCispubId { get; set; }
            public string OldSSDocnumber { get; set; }
            public string DocCNTLNo { get; set; }
            public string Invoice { get; set; }
            public string PONumber { get; set; }
            public string CispubProdId { get; set; }
            public string Ean { get; set; }
            public string ProductLookupKey { get; set; }
            public string Carrier { get; set; }
            public string FreightTerms { get; set; }
            public string MacFreightTerms { get; set; }
            public string PromoCode1 { get; set; }
            public string PromoCode2 { get; set; }
            public string Pageline { get; set; }
            public string DomesticforeignFlag { get; set; }
            public string PostFlags { get; set; }
            public string InitCode { get; set; }
            public string ESTACT { get; set; }
            public string DLY { get; set; }
            public string SSClaimNoCode { get; set; }
            public string BogusFlag { get; set; }
            public Nullable<System.DateTime> OrderEntryDate { get; set; }
            public Nullable<System.DateTime> ShipDate { get; set; }
            public Nullable<System.DateTime> InvoiceDate { get; set; }
            public Nullable<System.DateTime> PostDate { get; set; }
            public Nullable<System.DateTime> ReceiptDate { get; set; }
            public Nullable<int> GrossSalesUnits { get; set; }
            public Nullable<int> GrossRtnUnits { get; set; }
            public Nullable<int> InvNoPostQty { get; set; }
            public Nullable<int> RtnNoPostQty { get; set; }
            public Nullable<int> FiscalYear { get; set; }
            public Nullable<int> FiscalMonth { get; set; }
            public Nullable<int> SourceRowNumber { get; set; }
            public Nullable<decimal> GrossNativeCurrencyAmount { get; set; }
            public Nullable<decimal> USListPrice { get; set; }
            public Nullable<decimal> ListPrice { get; set; }
            public Nullable<decimal> SalesPrice { get; set; }
            public Nullable<decimal> CurrConvRate { get; set; }
            public Nullable<decimal> COGS { get; set; }
            public Nullable<decimal> DiscRate { get; set; }
            public Nullable<decimal> GrossSalesDols { get; set; }
            public Nullable<decimal> GrossRtnDols { get; set; }
            public Nullable<decimal> TransPortCost { get; set; }
            public Nullable<decimal> GrossCommDols { get; set; }
            public Nullable<decimal> GrossRoyDols { get; set; }
            public Nullable<decimal> RtnCommDols { get; set; }
            public Nullable<decimal> RtnRoyDols { get; set; }
            public Nullable<decimal> AltSalesDollars { get; set; }
            public Nullable<bool> IsException { get; set; }
            public string ExceptionNotes { get; set; }
            public Nullable<bool> IsApproved { get; set; }
            public Nullable<bool> ProcessedFlag { get; set; }
            public Nullable<int> CreatedBy { get; set; }
            public Nullable<System.DateTime> CreatedDate { get; set; }
            public Nullable<int> ModifiedBy { get; set; }
            public Nullable<System.DateTime> ModifiedDate { get; set; }
            public Nullable<bool> ActiveFlag { get; set; }
        }

    }
}
