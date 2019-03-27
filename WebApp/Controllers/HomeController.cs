using LinqToExcel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;
using System.Web.Mvc;
using WebApp.Models;
using Excel = Microsoft.Office.Interop.Excel;
namespace WebApp.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public ActionResult Index(FormCollection file)
        {
            string pathTest = "";
            try
            {
                using (var db = new AppDBContext())
                {
                    //  db.Configuration.AutoDetectChangesEnabled = false;
                    if (Request.Files["fiJPageInfo"].ContentLength > 0)
                    {
                        string fileExtension =
                                             System.IO.Path.GetExtension(Request.Files["fiJPageInfo"].FileName);
                        if (fileExtension == ".xls" || fileExtension == ".xlsx")
                        {
                            using (MemoryStream ms = new MemoryStream())
                            {
                                string fileLocation = Path.Combine(Server.MapPath("~/TempUpload/") + Path.GetFileName(Request.Files["fiJPageInfo"].FileName));
                                pathTest = fileLocation;
                                try
                                {
                                    DirectoryInfo di = new DirectoryInfo(Server.MapPath("~/TempUpload/"));
                                    if (!di.Exists)
                                    {
                                        di.Create();
                                    }
                                    if (System.IO.File.Exists(fileLocation))
                                    {
                                        System.IO.File.Delete(fileLocation);
                                    }
                                }
                                catch (Exception ex)
                                { }
                                Request.Files["fiJPageInfo"].SaveAs(fileLocation);

                                var excel = new ExcelQueryFactory();
                                excel.FileName = fileLocation;

                                List<EDI_754_945_ShipToChange> lstJpageNew = new List<EDI_754_945_ShipToChange>();
                                List<EDI_754_945_ShipToChange> lstJpageOld = new List<EDI_754_945_ShipToChange>();// db.EDI_DPCI_Data.ToList();
                                                                                            //Delete from table
                                db.Database.ExecuteSqlCommand("Exec SP_EDI_754_945_ShipToChange_Delete");

                                var details = (from x in excel.Worksheet<EDI_754_945_ShipToChange>(0) select x).ToList();

                                foreach (var detail in details)
                                {
                                    //EDI_754_945_ShipToChange objDetail = lstJpageOld.FirstOrDefault(x => x.CompanyCode == detail.CompanyCode &&
                                    //                                                          x.DivisionCode == detail.DivisionCode &&
                                    //                                                          x.CustomerNumber == detail.CustomerNumber &&
                                    //                                                          x.ITEM == detail.ITEM &&
                                    //                                                          x.PackCode == detail.PackCode &&
                                    //                                                          x.Style.Trim() == detail.Style.Trim());
                                    if (string.IsNullOrEmpty(detail.CompanyCode) || string.IsNullOrEmpty(detail.DivisionCode))
                                    {
                                        continue;
                                    }
                                    //if (objDetail == null)
                                    //{
                                    EDI_754_945_ShipToChange objDetail = new EDI_754_945_ShipToChange();
                                        objDetail.CompanyCode = detail.CompanyCode;
                                        objDetail.DivisionCode = detail.DivisionCode;
                                        objDetail.CustomerNumber = detail.CustomerNumber;
                                        objDetail.RRC_ = detail.RRC_;
                                        objDetail.Load_ID = detail.Load_ID;
                                        objDetail.SCAC = detail.SCAC;
                                        objDetail.Service_Level = detail.Service_Level;
                                        objDetail.Catalog_PO_Retail_DI = detail.Catalog_PO_Retail_DI;
                                        objDetail.Ship_Date = detail.Ship_Date;
                                        objDetail.Destination = detail.Destination;

                                        db.EDI_754_945_ShipToChange.Attach(objDetail);
                                        db.Entry(objDetail).State = System.Data.Entity.EntityState.Added;

                                        lstJpageNew.Add(objDetail);
                                    //}
                                    //else
                                    //{
                                    //    objDetail.Style = detail.Style;
                                    //    objDetail.Company = detail.Company;
                                    //    objDetail.ConversionQty = detail.ConversionQty;
                                    //    objDetail.Division = detail.Division;
                                    //    objDetail.ITEM = detail.ITEM;
                                    //    objDetail.Customer = detail.Customer;
                                    //    objDetail.UPC = detail.UPC;
                                    //    objDetail.PackCode = detail.PackCode;

                                    //    lstJpageNew.Add(objDetail);
                                    //    db.Entry(objDetail).State = System.Data.Entity.EntityState.Modified;
                                    //    db.EDI_DPCI_Data.Attach(objDetail);

                                    //}
                                }

                                db.SaveChanges();

                                ViewBag.Message = "File  successfully imported.";
                            }
                        }



                        // List<UploadResponseViewModel> lstSuccess = db.Database.SqlQuery<UploadResponseViewModel>("exec SP_Customer_XRef_EDI_Data").ToList();

                        //ViewBag.ListInserted = lstSuccess;// ds.Tables[0];
                        //ViewBag.ListError = lstError;

                        ViewBag.Message = "<div class='alert alert-warning'><span class='text-success'>items imported successfully,</span>";//<span class='text-error'> " + lstError.Count + " items has error. </span><button onclick='loadError()' class='btn btn-danger btn-sm error-btn'>View Error</button></div>";

                    }
                    else
                    {
                        ViewBag.ListInserted = null;
                        ViewBag.ListError = null;
                        ViewBag.Message = "<div class='alert alert-warning'>Please select valid file.</div>";

                    }
                    return View();
                }
            }
            catch (Exception ex)
            {
                ViewBag.Message = "Exception:" + ex.Message + " filepath:" + pathTest;
                return View();
            }
        }
        [AllowAnonymous]
        public ActionResult DownloadTemplate()
        {
            try
            {
                
                    string FilePath = Server.MapPath("~/754Template/");
                    DirectoryInfo dr = new DirectoryInfo(FilePath);
                    if (!dr.Exists)
                    {
                        dr.Create();
                    }
                    string FileName = "TemplateFileFrom754DataToUpdateAMT.xls";

                    byte[] fileBytes = System.IO.File.ReadAllBytes(FilePath + FileName);

                    return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, FileName);
                 
            }
            catch (Exception ex)
            {
                ViewBag.Message = ex.Message;
                return View();

            }

        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        private DataSet GetDataSet(string sql, CommandType commandType, Dictionary<string, Object> parameters)
        {
            // creates resulting dataset
            var result = new DataSet();

            // creates a data access context (DbContext descendant)
            using (var context = new AppDBContext())
            {
                // creates a Command 
                var cmd = context.Database.Connection.CreateCommand();
                cmd.CommandType = commandType;
                cmd.CommandText = sql;

                if (parameters != null && parameters.Count > 0)
                {
                    // adds all parameters
                    foreach (var pr in parameters)
                    {
                        var p = cmd.CreateParameter();
                        p.ParameterName = pr.Key;
                        p.Value = pr.Value;
                        cmd.Parameters.Add(p);
                    }
                }

                try
                {
                    // executes
                    context.Database.Connection.Open();
                    var reader = cmd.ExecuteReader();

                    // loop through all resultsets (considering that it's possible to have more than one)
                    do
                    {
                        // loads the DataTable (schema will be fetch automatically)
                        var tb = new DataTable();
                        tb.Load(reader);
                        result.Tables.Add(tb);

                    } while (!reader.IsClosed);
                }
                finally
                {
                    // closes the connection
                    context.Database.Connection.Close();
                }
            }

            // returns the DataSet
            return result;
        }
    }

    public class ExcelHelpers
    {
        /// <summary>
        /// Exports a list of objects to Excel. Objects go in the Rows while Object Properties go in the Columns
        /// </summary>
        /// <param name="objects">List of objects to export.</param>
        /// <param name="filePath">Location to save file to. Does not need to exist</param>
        /// <param name="fileName">Name of excel file</param>
        public static void ExportToExcel<T>(IEnumerable<T> objects, string filePath, string fileName)
        {
            // Add \ to end of file name if it doesn't exist. Just want to be consistant
            if (!filePath.EndsWith(@"\"))
                filePath += @"\";

            // Create directory if it doesn't exist
            if (!Directory.Exists(filePath))
                Directory.CreateDirectory(filePath);

            // Start Excel and get Application object. 
            Excel.Application excel = new Excel.Application();

            // Set it hidden and hide alerts
            excel.Visible = false;
            excel.DisplayAlerts = false;

            // Create a new workbook. 
            Excel.Workbook workbook = excel.Workbooks.Add();

            // Get the active sheet 
            Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;

            try
            {
                // Convert the list into a rectangular array that Excel can read
                var data = GetObjectArray<T>(objects);

                // If at least one record got converted successfully
                if (data.Length > 1)
                {
                    // Get the range of cells that the data will go into. Size matches rectangular array size
                    string xlsRange = string.Format("A1:{0}{1}",
                        new object[] { GetExcelColumn(data.GetLength(1)), data.GetLength(0) });

                    // Insert data into the specified range of cells
                    Excel.Range range = sheet.get_Range(xlsRange);
                    range.Value = data;

                    // Auto-Fit the columns
                    range.EntireColumn.AutoFit();
                }

                // Save workbook
                workbook.SaveAs(
                    string.Format("{0}{1}", new object[] { filePath, fileName }),
                    Excel.XlFileFormat.xlWorkbookDefault);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                // Close
                sheet = null;
                workbook.Close();
                workbook = null;
                excel.Quit();
            }

            // Clean up 
            // NOTE: When in release mode, this does the trick 
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        /// <summary>
        /// Takes a List of objects objects and converts the objects and their properties into a rectangular array of objects
        /// </summary>
        /// <param name="objects">List of objects to flatten</param>
        /// <returns>Rectangular array where objects are stored in [0] and properties are stored in [1]</returns>
        private static object[,] GetObjectArray<T>(IEnumerable<T> objects)
        {
            // Get list of object properties
            PropertyInfo[] properties = typeof(T).GetProperties();

            // Create rectangular array based on # of objects and # of object properties
            object[,] data = new object[objects.Count() + 1, properties.Length];

            // Loop through properties on object
            for (int j = 0; j < properties.Count(); j++)
            {
                // Write the property name into the first row of the array
                data[0, j] = properties[j].Name.Replace("_", " ");

                // Loop through objects and write out the specified property of each one into the array
                for (int i = 0; i < objects.Count(); i++)
                {
                    data[i + 1, j] = properties[j].GetValue(objects.ElementAt(i), null);
                }
            }

            // Return rectangular array
            return data;
        }

        /// <summary>
        /// Takes an Integer and converts it into Excel's column header code.
        /// For example: 1 = A; 2 = B; 27 = AA;
        /// </summary>
        /// <param name="colNumber">Number of Column in Excel. 1 = A</param>
        /// <returns>string that Excel can use</returns>
        private static string GetExcelColumn(int colNumber)
        {
            // If value is zero or less, return an empty string
            if (colNumber <= 0)
                return string.Empty;

            // If the value is less than or equal to 26 (Z), the column header
            // is only one character long. If it's greater, call this recursively
            // to get the first letter(s) of the column code.
            string first = (colNumber <= 26 ? string.Empty :
                GetExcelColumn((int)Math.Floor((colNumber - 1) / 26.00)));

            // Get the final letter in the column code
            int second = colNumber % 26;
            if (second == 0) second = 26;
            char finalLetter = (char)('A' + second - 1);            // Excel column header is the first part + the final character
            return string.Format("{0}{1}", new object[] { first, finalLetter });
        }
    }
}