using Microsoft.Office.Interop.Excel;
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
using System.Text;
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
        private static int excelRow;
        private static string baseURL;
        private static string keyFile;
        private static string ERROR_STATUS = "שגיאה";
        private static string KEY = String.Empty;
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
            baseURL = ConfigurationManager.AppSettings["baseUrl"];
            keyFile = ConfigurationManager.AppSettings["keyFile"];

            if (string.IsNullOrEmpty(uploadDir) || string.IsNullOrEmpty(archiveDir) || string.IsNullOrEmpty(excelFile) || 
                string.IsNullOrEmpty(CurrentBLine) || string.IsNullOrEmpty(baseURL) || string.IsNullOrEmpty(keyFile))
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


            KEY = GetKey();
            if (string.IsNullOrEmpty(KEY))
            {
                throw new Exception("בעיה בזמן נסיון לקבל את מחרוזת הרישיו - אנא בדוק את קובץ הלוג");
            }

            var upload = new DirectoryInfo(uploadDir);
            var clientDirs = upload.GetDirectories();


            
            SimpleLogger.SimpleLog.Info("available client count : " + clientDirs.Length);
            foreach (var clientDir in clientDirs)
            {
                var archive = GetArchiveDate();
                var adir = Path.Combine(archiveDir, archive);
                if (!Directory.Exists(adir))
                {
                    Directory.CreateDirectory(adir);
                }

                string bLineID = await GetBLine(clientDir);

                var workingDirs = clientDir.GetDirectories();
                SimpleLogger.SimpleLog.Info("available working directories count : " + workingDirs.Length);

                foreach (var workingDir1 in workingDirs)
                {
                    long date = 0;
                    DirectoryInfo workingDir = workingDir1;
                    var name = workingDir1.Name;
                    var splits = name.Split('_');
                    if (splits.Length == 2)
                    {
                        DateTime newdate;
                        var newName = splits[0];
                        if (DateTime.TryParseExact(splits[1], "ddMMyyyy", null, System.Globalization.DateTimeStyles.None, out newdate))
                        {
                            TimeSpan epochTicks = new TimeSpan(new DateTime(1970, 1, 1).Ticks);
                            TimeSpan unixTicks = new TimeSpan(newdate.Ticks) - epochTicks;
                            date = (long)unixTicks.TotalMilliseconds;
                        }
                        workingDir = new DirectoryInfo(Path.Combine(workingDir1.Parent.FullName, newName));
                        Directory.Move(workingDir1.FullName, workingDir.FullName);
                    }
                    var workingPath = workingDir.FullName;
                    SimpleLogger.SimpleLog.Info("workingPath folder : " + workingPath);
                    var calc = CalcDir(workingPath, clientDir.Name);
                    SimpleLogger.SimpleLog.Info("calc dir : " + calc.name + "--" + calc.docs);
                    excelRow = WriteToExcel(calc);
                    SimpleLogger.SimpleLog.Info("excel row - " + excelRow);
                    if (excelRow > 0)
                    {
                        SimpleLogger.SimpleLog.Info("after write to excel");

                        var info = await UploadToPortalAsync(workingPath, calc, bLineID, date);
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
                            var dest = Path.Combine(adir, clientDir.Name);
                            if (!Directory.Exists(dest))
                            {
                                Directory.CreateDirectory(dest);
                            }
                            dest = Path.Combine(dest, workingDir.Name);
                            Directory.Move(workingDir.FullName, dest);
                            SimpleLogger.SimpleLog.Info("after directory move to archive");
                        }
                    }
                }
            }
        }

        private static async Task<bool> UploadToPortalAsync(string dir, DirData calc, string bLineID, long date)
        {
            var caseId = await GetCaseID(calc.name, bLineID);
            SimpleLogger.SimpleLog.Info("in UploadToPortalAsync, case id : " + caseId);
            if (caseId == "ERROR")
            {
                return false;
            }
            if (caseId == null)
            {
                caseId = await CreateNewCase(calc.name, bLineID, date);
            }
            if (caseId == null)
            {
                SimpleLogger.SimpleLog.Info("Case is not uploaded. Case ID - " + calc.name);
                return false;
            }
            var uploads = await UploadFiles(calc.name, caseId, dir);
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
                        RequestUri = new Uri(baseURL + "/cases/" + caseId + "/process"),
                        Method = HttpMethod.Post,

                    };
                    client.DefaultRequestHeaders.Add("Authorization", "Bearer " + KEY);

                    using (var response = await client.SendAsync(request))
                    {
                        //response.EnsureSuccessStatusCode();
                        if (!response.IsSuccessStatusCode)
                        {
                            var body = await response.Content.ReadAsStringAsync();
                            HandleHttpError(response, body, request, name);
                        }
                    }
                    SimpleLogger.SimpleLog.Info("processed case . Case ID - " + caseId);
                    return true;
                }
                    
            }
            catch (Exception ex)
            {

                SimpleLogger.SimpleLog.Info("Error while starting process to case. Case ID - " + caseId + " ------- " + baseURL + "/cases/" + caseId + "/process");
                SimpleLogger.SimpleLog.Log(ex);
                BuildError(name, "Error while starting process to case. - " + ex.Message);
                return false;
            }
        }
        static async Task<string> GetBLine(DirectoryInfo dir)
        {
            string bLineID = bLines[CurrentBLine];

            var bLineName = dir.Name;
            Excel._Worksheet xlWorksheet = null;
            Excel.Workbook xlWorkbook = null;
            Excel.Application xlApp = null;
            var excelFileName = Path.GetFileName(excelFile);
            var bLineExcelFile = excelFile.Replace(excelFileName, "BusinessLine.csv");
            try
            {
                xlApp = new Excel.Application();
                xlApp.Visible = false;
                xlWorkbook = xlApp.Workbooks.Open(bLineExcelFile);
                xlWorksheet = (Excel._Worksheet)xlWorkbook.ActiveSheet;
                var lastRow = xlWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                for (int i = 2; i <= lastRow; i++)
                {
                    var name = xlWorksheet.Range["A" + i, "A" + i].Value2.ToString();
                    var bline = xlWorksheet.Range["B" + i, "B" + i].Value2.ToString();
                    if(name.ToLower() == bLineName.ToLower())
                    {
                        SimpleLogger.SimpleLog.Info("Trying to get bline from defaults. current bline = " + bLineID);
                        if (!bLines.TryGetValue(bline, out bLineID))
                        {
                            SimpleLogger.SimpleLog.Info("Getting Line from Owl");
                            bLineID = await GetBLineIdFromOwl(bline);
                            SimpleLogger.SimpleLog.Info("After getting bline from owl. current bline = " + bLineID);
                        }
                        if (bLineID == null || bLineID == "ERROR")
                        {
                            return bLines[CurrentBLine];
                        }
                    }
                }
               // xlWorkbook.Save();
                xlWorkbook.Close();
                xlApp.Quit();
                return bLineID;
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
            return null;
        }
        private static async Task<string> GetBLineIdFromOwl(string bline)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    var request = new HttpRequestMessage()
                    {
                        RequestUri = new Uri(baseURL + "/businessLines"),
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
        private static async Task<bool> UploadFiles(string name, string caseId, string xdir)
        {
            try
            {
                var dir = new DirectoryInfo(xdir);
                var files = dir.GetFiles().ToList();
                SimpleLogger.SimpleLog.Info("uploading " + files.Count() + " files from folder - " + dir.FullName + " to case id - " + caseId);
                for (int i = 0; i < files.Count(); i++)
                {
                   

                    var file = files[i];
                   
                    var fileExt = Path.GetExtension(file.FullName);
                    if(fileExt.ToLower() == ".ini")
                    {
                        continue;
                    }
                    var fileName = Path.GetFileNameWithoutExtension(file.FullName);
                    using (var client = new HttpClient())
                    {
                        FileInfo f = new FileInfo(file.FullName);
                        var sfile = new StreamContent(File.OpenRead(file.FullName));
                        sfile.Headers.ContentType = new MediaTypeHeaderValue(GetMimeType(Path.GetExtension(file.FullName)));
                        sfile.Headers.ContentLength = f.Length;
                        client.DefaultRequestHeaders.Add("Authorization", "Bearer " + KEY);
                        client.DefaultRequestHeaders.Add("x-case-id", caseId);
                        client.DefaultRequestHeaders.Add("x-file-name", fileName);
                        using (var post = await client.PostAsync(baseURL + "/documents", sfile))
                        {
                            //post.EnsureSuccessStatusCode();
                            if (!post.IsSuccessStatusCode)
                            {
                                var body = await post.Content.ReadAsStringAsync();
                                HandleHttpError(post, body, null, name);
                            }
                        }
                    }
                        
                   

                }
                SimpleLogger.SimpleLog.Info("uploaded " + files.Count() + " files to case id - " + caseId);
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
        public static string GetMimeType(string fileExtension)
        {
            // Normalize the extension by removing the dot if present and converting to lowercase
            fileExtension = fileExtension.TrimStart('.').ToLower();

            switch (fileExtension)
            {
                case "jpeg":
                    return "image/jpeg";
                case "jpg":
                    return "image/jpeg";
                case "bmp":
                    return "image/bmp";
                case "pdf":
                    return "application/pdf";
                case "doc":
                    return "application/msword";
                case "docx":
                    return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                case "tif":
                    return "image/tiff";
                case "tiff":
                    return "image/tiff";
                default:
                    return "application/pdf";
            }
        }
        private static async Task<string> GetCaseID(string name, string bLineID)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    var request = new HttpRequestMessage()
                    {
                        RequestUri = new Uri(baseURL + "/cases?search=" + name),
                        Method = HttpMethod.Get,

                    };
                    client.DefaultRequestHeaders.Add("Authorization", "Bearer " + KEY);

                    using (var response = await client.SendAsync(request))
                    {
                        //response.EnsureSuccessStatusCode();
                        if (!response.IsSuccessStatusCode)
                        {
                            var body = await response.Content.ReadAsStringAsync();
                            HandleHttpError(response, body, request, name);
                            return "ERROR";
                        }
                        
                        var status = response.StatusCode;
                        var data = await response.Content.ReadAsStringAsync();
                        var oData = (JArray)JsonConvert.DeserializeObject(data);
                        if (oData.Count == 0)
                        {
                            return null;
                        }
                        var obj = oData.Children<JObject>().FirstOrDefault(f => f["name"] != null && f["name"].ToString() == name && f["businessLineId"].ToString() == bLineID);
                        if (obj != null && obj.Count > 0)
                        {
                            var cstatus = obj["externalStatus"]?.ToString();
                            var id = obj["id"].ToString();
                            if (cstatus == "archived")
                            {
                                await unArchiveCase(id, name);
                            }
                            return id;
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

        private async static Task<bool> unArchiveCase(string caseId, string name)
        {
            try
            {
                SimpleLogger.SimpleLog.Info("unArchive a case . Case ID - " + caseId);
                using (var client = new HttpClient())
                {
                    var request = new HttpRequestMessage()
                    {
                        RequestUri = new Uri(baseURL + "/cases/" + caseId + "/unarchive"),
                        Method = HttpMethod.Put,

                    };
                    client.DefaultRequestHeaders.Add("Authorization", "Bearer " + KEY);

                    using (var response = await client.SendAsync(request))
                    {
                        //response.EnsureSuccessStatusCode();
                        
                        if (!response.IsSuccessStatusCode)
                        {
                            var body = await response.Content.ReadAsStringAsync();
                            HandleHttpError(response, body, request, name);
                        }
                    }
                    SimpleLogger.SimpleLog.Info("unArchive a case . Case ID - " + caseId);
                    return true;
                }

            }
            catch (Exception ex)
            {

                SimpleLogger.SimpleLog.Info("Error while unArchive a case. Case ID - " + caseId + " ------- " + baseURL + "/cases/" + caseId + "/process");
                SimpleLogger.SimpleLog.Log(ex);
                BuildError(name, "Error while unArchive a case. - " + ex.Message);
                return false;
            }
        }

        //private static async Task<bool> CheckCaseExist(string name)
        //{
        //    try
        //    {
        //        var client = new HttpClient();
        //        var request = new HttpRequestMessage()
        //        {
        //            RequestUri = new Uri(baseURL + "/cases/" + name),
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
        private static async Task<string> CreateNewCase(string name, string bLineID, long date)
        {
            try
            {
                var client = new HttpClient();
                var request = new HttpRequestMessage()
                {
                    RequestUri = new Uri(baseURL + "/cases"),
                    Method = HttpMethod.Post,
                };
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + KEY);
                var currentCase = new Case
                {
                    name = name,
                    businessLineId = bLineID
                };
                if (date > 0)
                {
                    currentCase.keyDecisionDate = date;
                }
                var json = JsonConvert.SerializeObject(currentCase);

                HttpContent _Body = new StringContent(json);
                _Body.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                request.Content = _Body;



                
                using (var response = await client.SendAsync(request))
                {
                    //response.EnsureSuccessStatusCode();
                    var body = await response.Content.ReadAsStringAsync();
                    if (!response.IsSuccessStatusCode)
                    {
                        HandleHttpError(response, body, request, name);
                        return null;
                    }
                    var obj = (JObject)JsonConvert.DeserializeObject(body);
                    return obj["id"].ToString();

                }
                
                
            }
            
            catch (Exception ex)
            {
                SimpleLogger.SimpleLog.Info("Failed to create a case with DigitalOwl. Case ID - " + name);
                SimpleLogger.SimpleLog.Log(ex);
                BuildError(name, "Failed to create a case with DigitalOwl. - " + ex.Message);
                return null;
            }
            
        }
        private static void HandleHttpError(HttpResponseMessage response, string body, HttpRequestMessage request, string name)
        {
            try
            {
                var errorMessage = new StringBuilder();
                // Status code and reason phrase
                var statusCode = (int)response.StatusCode;
                var reasonPhrase = response.ReasonPhrase;

                // Response headers
                var responseHeaders = response.Headers;
                var contentHeaders = response.Content.Headers;

                // Attempt to parse the response body as JSON
                string errorDetails = body;
                try
                {
                    var errorObject = JObject.Parse(body);
                    errorDetails = errorObject.ToString();
                }
                catch
                {
                    // If parsing fails, keep the raw body
                }

                // Build a detailed error message

                errorMessage.AppendLine($"Request failed with status code {statusCode} ({reasonPhrase}).");
                if(request != null)
                {
                    errorMessage.AppendLine($"Request Method: {request.Method}");
                    errorMessage.AppendLine($"Request URI: {request.RequestUri}");

                    // Include request headers
                    errorMessage.AppendLine("Request Headers:");
                    foreach (var header in request.Headers)
                    {
                        errorMessage.AppendLine($"{header.Key}: {string.Join(", ", header.Value)}");
                    }
                }
                

                // Include response headers
                errorMessage.AppendLine("Response Headers:");
                foreach (var header in responseHeaders)
                {
                    errorMessage.AppendLine($"{header.Key}: {string.Join(", ", header.Value)}");
                }

                // Include content headers
                errorMessage.AppendLine("Content Headers:");
                foreach (var header in contentHeaders)
                {
                    errorMessage.AppendLine($"{header.Key}: {string.Join(", ", header.Value)}");
                }

                // Include the response content
                errorMessage.AppendLine("Response Content:");
                errorMessage.AppendLine(errorDetails);

                // Throw an exception with the detailed error message
                SimpleLogger.SimpleLog.Info("Http error when trying to create a case with DigitalOwl. Case ID - " + name);
                SimpleLogger.SimpleLog.Error(errorMessage.ToString());
                BuildError(name, "Http Error trying to create a case with DigitalOwl. - " + errorDetails);
            }
            catch (Exception ex)
            {
                SimpleLogger.SimpleLog.Info("Error with HandleHttpError method");
                SimpleLogger.SimpleLog.Log(ex);
                BuildError(name, "Failed to create a case with DigitalOwl. - " + ex.Message);
            }
            
        }
        static DirData CalcDir(string adir, string bline)
        {
            var dir = new DirectoryInfo(adir);
            var files = dir.GetFiles("*.pdf", SearchOption.TopDirectoryOnly);
            var count = files.Count();
            return new DirData
            {
                date = FormatExcelDate(DateTime.Now),
                name = dir.Name,
                docs = count.ToString(),
                status = "רישום",
                bline = bline
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
                xlWorksheet.Range[E_REMARK + excelRow, E_REMARK + excelRow].Value2 = data.date;
                xlWorksheet.Range[E_STATUS + excelRow, E_STATUS + excelRow].Value2 = data.status;
                //var lastRow = xlWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;

                //for (int i = lastRow; i > 1; i--)
                //{
                //    var name = xlWorksheet.Range[E_NAME + i, E_NAME + i].Value2.ToString();
                //    var status = xlWorksheet.Range[E_STATUS + i, E_STATUS + i].Value2.ToString();
                //    if (name == data.name && status == "רישום")
                //    {
                //        xlWorksheet.Range[E_REMARK + i, E_REMARK + i].Value2 = data.date;
                //        xlWorksheet.Range[E_STATUS + i, E_STATUS + i].Value2 = data.status;
                //    }
                //}



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

                SimpleLogger.SimpleLog.Info("updating excel status for row - " + excelRow);

                xlWorksheet.Range[E_DATEUPLOAD + excelRow, E_DATEUPLOAD + excelRow].Value2 = data.date;
                xlWorksheet.Range[E_STATUS + excelRow, E_STATUS + excelRow].Value2 = data.status;


                //var lastRow = xlWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;

                //for (int i = lastRow; i > 1; i--)
                //{
                //    var name = xlWorksheet.Range[E_NAME + i, E_NAME + i].Value2.ToString();
                //    var status = xlWorksheet.Range[E_STATUS + i, E_STATUS + i].Value2.ToString();
                //    if (name == data.name && status == "רישום")
                //    {
                //        xlWorksheet.Range[E_DATEUPLOAD + i, E_DATEUPLOAD + i].Value2 = data.date;
                //        xlWorksheet.Range[E_STATUS + i, E_STATUS + i].Value2 = data.status;
                //    }
                //}


               
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
        static int WriteToExcel(DirData data)
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
                for (int i = 2; i <= lastRow; i++)
                {
                    var status = xlWorksheet.Range[E_STATUS + i, E_STATUS + i].Value2.ToString().Trim();
                    var name = xlWorksheet.Range[E_NAME + i, E_NAME + i].Value2.ToString().Trim();
                    if(name == data?.name && status == ERROR_STATUS)
                    {
                        row = i;
                    }
                }
                

                SimpleLogger.SimpleLog.Info("before write to excel property loop");
                foreach (PropertyInfo prop in data.GetType().GetProperties())
                {
                    switch (prop.Name)
                    {
                        case "date":
                            xlWorksheet.Range[E_DATE + row, E_DATE + row].Value2 = data?.date;
                            break;
                        case "name":
                            xlWorksheet.Range[E_NAME + row, E_NAME + row].Value2 = data?.name.Trim();
                            break;
                        case "docs":
                            xlWorksheet.Range[E_NUMDOCS + row, E_NUMDOCS + row].Value2 = data?.docs;
                            break;
                        case "status":
                            xlWorksheet.Range[E_STATUS + row, E_STATUS + row].Value2 = data?.status;
                            break;
                        case "bline":
                            xlWorksheet.Range[E_BLINE + row, E_BLINE + row].Value2 = data?.bline.Trim();
                            break;
                    }
                }
                SimpleLogger.SimpleLog.Info("after write to excel property loop");
                xlWorkbook.Save();
                xlWorkbook.Close();
                xlApp.Quit();
                return row;

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
            return -1;
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
                status = ERROR_STATUS
            };
            ErrorToExcel(errorStatus);
        }

        private static string GetKey()
        {
            Microsoft.Office.Interop.Word.Application word = null;
            Microsoft.Office.Interop.Word.Document doc = null;

            try
            {
                string key = string.Empty;
                word = new Microsoft.Office.Interop.Word.Application();
                doc = word.Documents.Open(keyFile);
                foreach (Microsoft.Office.Interop.Word.Paragraph objParagraph in doc.Paragraphs)
                {
                    key = objParagraph.Range.Text.Trim();
                }
                doc.Close();
                word.Quit();
                return key;
            }
            catch (Exception ex)
            {
                SimpleLogger.SimpleLog.Info("Error while trying to get the license key from file - " + ex.Message);
                SimpleLogger.SimpleLog.Log(ex);
                return string.Empty;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                if (doc != null)
                {
                    Marshal.ReleaseComObject(doc);
                }

                if (word != null)
                {
                    Marshal.ReleaseComObject(word);
                }
            }
        }


        private static string E_DATE = "A";
        private static string E_NAME = "B";
        private static string E_NUMDOCS = "C";
        private static string E_NUMPAGES = "D";
        private static string E_MEDICALDATA = "E";
        private static string E_HANDWRITTEN = "F";
        private static string E_STATUS = "G";
        private static string E_OWLSTATUS = "H";
        private static string E_DATEUPLOAD = "I";
        private static string E_DATEDOWNLOAD = "J";
        private static string E_CONTINUE = "K";
        private static string E_BLINE = "L";
        private static string E_REMARK = "M";
        
    }
    public class DirData
    {
        public string date { get; set; }
        public string name { get; set; }
        public string docs { get; set; }
        public string status { get; set; }
        public string remark { get; set; }
        public string bline { get; set; }

    }
    public class Case
    {
        public string name { get; set; }
        public string businessLineId { get; set; }
        public long keyDecisionDate { get; set; }
    }
}
