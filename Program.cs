using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DigtalOwl_Upload
{
    class Program
    {
        private static string uploadDir;
        private static string archiveDir;
        private static string excelFile;
        private static string CurrentBLine;
        private static string KEY = "eyJjbGllbnRfaWQiOiI0dzFPNUlJTE9GajdSajhvckZqTkJvR3Z4RVkwNlhUQyIsImNsaWVudF9zZWNyZXQiOiJsYWNBcFd0VTFHeXRfSVNlVGZCZVdweGRBRVJ3NG94Zm9EWkNvZmw0NjI2N3p1Q3ZSRUFTRjdpSEFDWDRnSmIzIiwiYXVkaWVuY2UiOiJodHRwczovL2FwaS5kaWdpdGFsb3dsLmFwcCIsImdyYW50X3R5cGUiOiJjbGllbnRfY3JlZGVudGlhbHMifQ==";
        private static Dictionary<string, string> bLines = new Dictionary<string, string>
        {
            {"defBlName", "914aa316-2243-4efb-aeea-a61758772b38" },
            {"defBlTemp", "9ff4ab50-58ee-4f3e-9c95-7479c6e02529"}
        };

        
        static async Task Main(string[] args)
        {
            uploadDir = ConfigurationManager.AppSettings["uploadDir"];
            archiveDir = ConfigurationManager.AppSettings["archivedDir"];
            excelFile = ConfigurationManager.AppSettings["excelFile"];
            CurrentBLine = ConfigurationManager.AppSettings["buisnessLine"];

            if (string.IsNullOrEmpty(uploadDir) || string.IsNullOrEmpty(archiveDir) || string.IsNullOrEmpty(excelFile) || string.IsNullOrEmpty(CurrentBLine))
            {
                throw new Exception("פרטי קונפיגורציה חסרים");
            }
            if (!Directory.Exists(uploadDir))
            {
                Directory.CreateDirectory(uploadDir);
            }
            if (!Directory.Exists(archiveDir))
            {
                Directory.CreateDirectory(archiveDir);
            }
            var upload = new DirectoryInfo(uploadDir);
            var udirs = upload.GetDirectories();


            string bLineID;
            if (!bLines.TryGetValue(CurrentBLine, out bLineID))
            {
                bLineID = await GetBLineID(CurrentBLine);
            }
            if (bLineID == null || bLineID == "ERROR")
            {
                SimpleLogger.SimpleLog.Info("No BuisnessLine exist for provided value");
                throw new Exception("No BuisnessLine exist for provided value");
            }
            SimpleLogger.SimpleLog.Info("available directories count : " + udirs.Length);
            foreach (var udir in udirs)
            {
                var archive = GetArchiveDate();
                var adir = Path.Combine(archiveDir, archive);
                if (!Directory.Exists(adir))
                {
                    Directory.CreateDirectory(adir);
                }

                var workingPath = udir.FullName;
                SimpleLogger.SimpleLog.Info("workingPath folder : " + workingPath); 
                var calc = CalcDir(workingPath);
                SimpleLogger.SimpleLog.Info("calc dir : " + calc.name + "--" + calc.docs); 
                if (WriteToExcel(calc))
                {
                    SimpleLogger.SimpleLog.Info("after write to excel");
                    var info = await UploadToPortalAsync(workingPath, calc, bLineID);
                    if (info)
                    {
                        SimpleLogger.SimpleLog.Info("after upload");
                        var newStatus = new DirData
                        {
                            date = FormatExcelDate(DateTime.Now),
                            name = calc.name,
                            status = "העלה"
                        };
                        UpdateExcelStatus(newStatus);
                        SimpleLogger.SimpleLog.Info("after update status");
                        var dest = Path.Combine(adir, udir.Name);
                        Directory.Move(udir.FullName, dest);
                        SimpleLogger.SimpleLog.Info("after directory move to archive");
                    }
                }
                
            }
        }

        private static async Task<bool> UploadToPortalAsync(string dest, DirData calc, string bLineID)
        {
            var caseId = await GetCaseID(calc.name);
            SimpleLogger.SimpleLog.Info("in UploadToPortalAsync, case id : " + caseId);
            if (caseId == "ERROR")
            {
                return false;
            }
            if (caseId == null)
            {
                caseId = await CreateNewCase(calc.name, bLineID);
            }
            if (caseId == null)
            {
                SimpleLogger.SimpleLog.Info("Case is not uploaded. Case ID - " + calc.name);
                return false;
            }
            var uploads = await UploadFiles(calc.name, caseId, dest);
            if (!uploads)
            {
                SimpleLogger.SimpleLog.Info("Failed to upload files to case. Case ID - " + calc.name);
                return false;
            }
            var process = await ProcessFile(calc.name, caseId);
            if (!process)
            {
                return false;
            }
            return true;
        }
        private static async Task<bool> ProcessFile(string name, string caseId)
        {
            try
            {
                SimpleLogger.SimpleLog.Info("processing case . Case ID - " + caseId);
                using (var client = new HttpClient())
                {
                    var request = new HttpRequestMessage()
                    {
                        RequestUri = new Uri("https://api.digitalowl.app/cases/" + caseId + "/process"),
                        Method = HttpMethod.Post,

                    };
                    client.DefaultRequestHeaders.Add("Authorization", "Bearer " + KEY);

                    using (var response = await client.SendAsync(request))
                    {
                        response.EnsureSuccessStatusCode();
                    }
                    SimpleLogger.SimpleLog.Info("processed case . Case ID - " + caseId);
                    return true;
                }
                    
            }
            catch (Exception ex)
            {

                SimpleLogger.SimpleLog.Info("Error while starting process to case. Case ID - " + caseId + " ------- " + "https://api.digitalowl.app/cases/" + caseId + "/process");
                SimpleLogger.SimpleLog.Log(ex);
                BuildError(name, "Error while starting process to case. - " + ex.Message);
                return false;
            }
        }
        private static async Task<string> GetBLineID(string bline)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    var request = new HttpRequestMessage()
                    {
                        RequestUri = new Uri("https://api.digitalowl.app/businessLines"),
                        Method = HttpMethod.Get,

                    };
                    client.DefaultRequestHeaders.Add("Authorization", "Bearer " + KEY);

                    using (var response = await client.SendAsync(request))
                    {
                        var data = await response.Content.ReadAsStringAsync();
                        var oData = (JArray)JsonConvert.DeserializeObject(data);
                        if (oData.Count == 0)
                        {
                            return null;
                        }
                        var obj = oData.Children<JObject>().FirstOrDefault(f => f["name"] != null && f["name"].ToString() == bline);
                        if (obj != null && obj.Count > 0)
                        {
                            return obj["id"].ToString();

                        }
                        return null;
                    }
                }
                    
            }
            catch (Exception ex)
            {
                SimpleLogger.SimpleLog.Info("Error while checking id for buisness line ID - " + bline);
                SimpleLogger.SimpleLog.Log(ex);
                return "ERROR";
            }
        }
        private static async Task<bool> UploadFiles(string name, string caseId, string dest)
        {
            try
            {
                var dir = new DirectoryInfo(dest);
                var files = dir.GetFiles().ToList();
                for (int i = 0; i < files.Count(); i++)
                {
                    var file = files[i];
                    var fileName = Path.GetFileNameWithoutExtension(file.FullName);
                    using (var client = new HttpClient())
                    {
                        FileInfo f = new FileInfo(file.FullName);
                        var sfile = new StreamContent(File.OpenRead(file.FullName));
                        sfile.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
                        sfile.Headers.ContentLength = f.Length;
                        client.DefaultRequestHeaders.Add("Authorization", "Bearer " + KEY);
                        client.DefaultRequestHeaders.Add("x-case-id", caseId);
                        client.DefaultRequestHeaders.Add("x-file-name", fileName);
                        using (var post = await client.PostAsync("https://api.digitalowl.app/documents", sfile))
                        {
                            post.EnsureSuccessStatusCode();
                        }
                    }
                        
                   

                }
                return true;
            }
            catch (Exception ex)
            {
                SimpleLogger.SimpleLog.Info("Failed to upload documents. Case ID - " + name);
                SimpleLogger.SimpleLog.Log(ex);
                BuildError(name, "Failed to upload documents. - " + ex.Message);
                return false;
            }
        }
        private static async Task<string> GetCaseID(string name)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    var request = new HttpRequestMessage()
                    {
                        RequestUri = new Uri("https://api.digitalowl.app/cases?search=" + name),
                        Method = HttpMethod.Get,

                    };
                    client.DefaultRequestHeaders.Add("Authorization", "Bearer " + KEY);

                    using (var response = await client.SendAsync(request))
                    {
                        var status = response.StatusCode;
                        var data = await response.Content.ReadAsStringAsync();
                        var oData = (JArray)JsonConvert.DeserializeObject(data);
                        if (oData.Count == 0)
                        {
                            return null;
                        }
                        var obj = oData.Children<JObject>().FirstOrDefault(f => f["name"] != null && f["name"].ToString() == name);
                        if (obj != null && obj.Count > 0)
                        {
                            return obj["id"].ToString();
                        }
                        return null;
                    }
                }
                    
            }
            catch (Exception ex)
            {
                SimpleLogger.SimpleLog.Info("Error while checking is case exists. Case ID - " + name);
                SimpleLogger.SimpleLog.Log(ex);
                BuildError(name, "Error while checking is case exists. - " + ex.Message);
                return "ERROR";
            }
        }
        //private static async Task<bool> CheckCaseExist(string name)
        //{
        //    try
        //    {
        //        var client = new HttpClient();
        //        var request = new HttpRequestMessage()
        //        {
        //            RequestUri = new Uri("https://api.digitalowl.app/cases/" + name),
        //            Method = HttpMethod.Get,
                    
        //        };
        //        client.DefaultRequestHeaders.ExpectContinue = false;
        //        client.DefaultRequestHeaders.Add("Authorization", "Bearer " + KEY);
                
        //        using (var response = await client.SendAsync(request))
        //        {
        //            var status = response.StatusCode;
        //            return status == HttpStatusCode.OK;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        SimpleLogger.SimpleLog.Info("Error while checking is case exists. Case ID - " + name);
        //        SimpleLogger.SimpleLog.Log(ex);
        //        return false;
        //    }


        //}
        private static async Task<string> CreateNewCase(string name, string bLineID)
        {
            try
            {
                var client = new HttpClient();
                var request = new HttpRequestMessage()
                {
                    RequestUri = new Uri("https://api.digitalowl.app/cases"),
                    Method = HttpMethod.Post,
                };
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + KEY);
                var json = JsonConvert.SerializeObject(new Case
                {
                    name = name,
                    businessLineId = bLineID

                });

                HttpContent _Body = new StringContent(json);
                _Body.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                request.Content = _Body;




                using (var response = await client.SendAsync(request))
                {
                    response.EnsureSuccessStatusCode();
                    var body = await response.Content.ReadAsStringAsync();
                    var obj = (JObject)JsonConvert.DeserializeObject(body);
                    return obj["id"].ToString();
                }
            }
            catch(Exception ex)
            {
                SimpleLogger.SimpleLog.Info("Failed to create a case with DigitalOwl. Case ID - " + name);
                SimpleLogger.SimpleLog.Log(ex);
                BuildError(name, "Failed to create a case with DigitalOwl. - " + ex.Message);
                return null;
            }
            
        }

        static DirData CalcDir(string adir)
        {
            var dir = new DirectoryInfo(adir);
            var files = dir.GetFiles("*.pdf", SearchOption.TopDirectoryOnly);
            var count = files.Count();
            return new DirData
            {
                date = FormatExcelDate(DateTime.Now),
                name = dir.Name,
                docs = count.ToString(),
                status = "רישום"
            };

        }
        static void ErrorToExcel(DirData data)
        {
            Excel._Worksheet xlWorksheet = null;
            Excel.Workbook xlWorkbook = null;
            Excel.Application xlApp = null;
            try
            {
                xlApp = new Excel.Application();
                xlApp.Visible = false;
                xlWorkbook = xlApp.Workbooks.Open(excelFile);
                xlWorksheet = (Excel._Worksheet)xlWorkbook.ActiveSheet;
                var lastRow = xlWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;

                for (int i = lastRow; i > 1; i--)
                {
                    var name = xlWorksheet.Range[E_NAME + i, E_NAME + i].Value2.ToString();
                    var status = xlWorksheet.Range[E_STATUS + i, E_STATUS + i].Value2.ToString();
                    if (name == data.name && status == "רישום")
                    {
                        xlWorksheet.Range[E_REMARK + i, E_REMARK + i].Value2 = data.date;
                        xlWorksheet.Range[E_STATUS + i, E_STATUS + i].Value2 = data.status;
                    }
                }



                xlWorkbook.Save();
                xlWorkbook.Close();
                xlApp.Quit();
                //

            }
            catch (Exception ex)
            {
                xlWorkbook.Close();
                xlApp.Quit();
                SimpleLogger.SimpleLog.Log(ex);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                if (xlWorksheet != null)
                {
                    Marshal.ReleaseComObject(xlWorksheet);
                }
                //close and release
                if (xlWorkbook != null)
                {
                    Marshal.ReleaseComObject(xlWorkbook);
                }

                if (xlApp != null)
                {
                    //quit and release

                    Marshal.ReleaseComObject(xlApp);
                }
            }
        }
        static void UpdateExcelStatus(DirData data)
        {
            Excel._Worksheet xlWorksheet = null;
            Excel.Workbook xlWorkbook = null;
            Excel.Application xlApp = null;
            try
            {
                xlApp = new Excel.Application();
                xlApp.Visible = false;
                xlWorkbook = xlApp.Workbooks.Open(excelFile);
                xlWorksheet = (Excel._Worksheet)xlWorkbook.ActiveSheet;
                var lastRow = xlWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;

                for (int i = lastRow; i > 1; i--)
                {
                    var name = xlWorksheet.Range[E_NAME + i, E_NAME + i].Value2.ToString();
                    var status = xlWorksheet.Range[E_STATUS + i, E_STATUS + i].Value2.ToString();
                    if (name == data.name && status == "רישום")
                    {
                        xlWorksheet.Range[E_DATEUPLOAD + i, E_DATEUPLOAD + i].Value2 = data.date;
                        xlWorksheet.Range[E_STATUS + i, E_STATUS + i].Value2 = data.status;
                    }
                }


               
                xlWorkbook.Save();
                xlWorkbook.Close();
                xlApp.Quit();
                //

            }
            catch (Exception ex)
            {
                xlWorkbook.Close();
                xlApp.Quit();
                SimpleLogger.SimpleLog.Log(ex);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                if (xlWorksheet != null)
                {
                    Marshal.ReleaseComObject(xlWorksheet);
                }
                //close and release
                if (xlWorkbook != null)
                {
                    Marshal.ReleaseComObject(xlWorkbook);
                }

                if (xlApp != null)
                {
                    //quit and release

                    Marshal.ReleaseComObject(xlApp);
                }
            }
        }
        static bool WriteToExcel(DirData data)
        {
            Excel._Worksheet xlWorksheet = null;
            Excel.Workbook xlWorkbook = null;
            Excel.Application xlApp = null;
            try
            {
                xlApp = new Excel.Application();
                xlApp.Visible = false;
                xlWorkbook = xlApp.Workbooks.Open(excelFile);
                xlWorksheet = (Excel._Worksheet)xlWorkbook.ActiveSheet;
                var lastRow = xlWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                var row = lastRow + 1;

                SimpleLogger.SimpleLog.Info("before write to excel property loop");
                foreach (PropertyInfo prop in data.GetType().GetProperties())
                {
                    switch (prop.Name)
                    {
                        case "date":
                            xlWorksheet.Range[E_DATE + row, E_DATE + row].Value2 = data?.date;
                            break;
                        case "name":
                            xlWorksheet.Range[E_NAME + row, E_NAME + row].Value2 = data?.name;
                            break;
                        case "docs":
                            xlWorksheet.Range[E_NUMDOCS + row, E_NUMDOCS + row].Value2 = data?.docs;
                            break;
                        case "status":
                            xlWorksheet.Range[E_STATUS + row, E_STATUS + row].Value2 = data?.status;
                            break;
                    }
                }
                SimpleLogger.SimpleLog.Info("after write to excel property loop");
                xlWorkbook.Save();
                xlWorkbook.Close();
                xlApp.Quit();
                return true;

            }
            catch (Exception ex)
            {
                SimpleLogger.SimpleLog.Log(ex);
                xlWorkbook.Close();
                xlApp.Quit();
                
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                if (xlWorksheet != null)
                {
                    Marshal.ReleaseComObject(xlWorksheet);
                }
                //close and release
                if (xlWorkbook != null)
                {
                    Marshal.ReleaseComObject(xlWorkbook);
                }

                if (xlApp != null)
                {
                    //quit and release

                    Marshal.ReleaseComObject(xlApp);
                }
            }
            return false;
        }
        static string FormatExcelDate(DateTime dt)
        {
            return dt.ToString("dd/MM/yyyy H:mm:ss");
        }
        static string GetArchiveDate() {
            return DateTime.Now.ToString("yyyyMMddHH");
        }
        private static void BuildError(string name, string msg)
        {
            var errorStatus = new DirData
            {
                remark = msg,
                name = name,
                status = "שגיאה"
            };
            ErrorToExcel(errorStatus);
        }
        private static string E_DATE = "A";
        private static string E_NAME = "B";
        private static string E_NUMDOCS = "C";
        private static string E_NUMPAGES = "D";
        private static string E_STATUS = "E";
        private static string E_DATEUPLOAD = "F";
        private static string E_DATEDOWNLOAD = "G";
        private static string E_CONTINUE = "H";
        private static string E_REMARK = "I";
    }
    public class DirData
    {
        public string date { get; set; }
        public string name { get; set; }
        public string docs { get; set; }
        public string status { get; set; }
        public string remark { get; set; }


    }
    public class Case
    {
        public string name { get; set; }
        public string businessLineId { get; set; }
    }
}
