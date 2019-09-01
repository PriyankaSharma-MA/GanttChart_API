using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using GanttChart.Models;
using System.IO;
using System.Web;
using System.Configuration;
using System.Data.OleDb;
using System.Data;
using System.Linq;
using System.Drawing;


using System.Text;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using System.Diagnostics;
using GanttChart.App_Code;

namespace GanttChart.Controllers
{
    public class CSVController : ApiController
    {
        string CSVPath = System.Configuration.ConfigurationManager.AppSettings["CSVPath"];
        string SharepointPath = System.Configuration.ConfigurationManager.AppSettings["SharepointPath"];
        string Sharepointfilename = System.Configuration.ConfigurationManager.AppSettings["Sharepointfilename"];
       // string Username = System.Configuration.ConfigurationManager.AppSettings["UserName"];
      //  string Password = System.Configuration.ConfigurationManager.AppSettings["Password"];
     

        [HttpGet]
        [ActionName("getAllExcelData")]
        public List<ExcelData> getAllExcelData([FromUri] string filename)
        {
            string conStr = "", Extension = ".xlsx";
          //  uploadSharePointFileFunction();
            switch (Extension)
            {
                case ".xls": //Excel 97-03
                    conStr = ConfigurationManager.ConnectionStrings["Excel03ConString"]
                             .ConnectionString;
                    break;
                case ".xlsx": //Excel 07
                    conStr = ConfigurationManager.ConnectionStrings["Excel07ConString"]
                              .ConnectionString;
                    break;
            }
            conStr = String.Format(conStr, Path.Combine(HttpContext.Current.Server.MapPath("~/CSV/CurrentFile"), filename));
            OleDbConnection connExcel = new OleDbConnection(conStr);

            string Roadmapsheet = "Official Roadmap$";
            string Roadmapcolor = "Color$";
            string Roadmapresource = "ProgramOverview$";

            ExcelRoadMapdata excelRoadMapdata;
            ExcelColordata excelColordata;
            ExcelResourcedata excelResourcedata;

            ExcelData excelData;
            List<ExcelRoadMapdata> lstExcelRoadMapdata = new List<ExcelRoadMapdata>();
            List<ExcelColordata> lstExcelColordata = new List<ExcelColordata>();
            List<ExcelResourcedata> lstExcelResourcedata = new List<ExcelResourcedata>();
            List<ExcelData> lstExcelData = new List<ExcelData>();

            try
            {

                OleDbCommand cmdExcel = new OleDbCommand();
                OleDbDataAdapter oda = new OleDbDataAdapter();
                System.Data.DataTable dt = new System.Data.DataTable();
                System.Data.DataTable dtRoadmap = new System.Data.DataTable();
                System.Data.DataTable dtColor = new System.Data.DataTable();
                System.Data.DataTable dtResource = new System.Data.DataTable();
                cmdExcel.Connection = connExcel;

                //Get the name of First Sheet              


                //Read Data from First Sheet
                connExcel.Open();
                // cmdExcel.CommandText = "SELECT * From [" + Roadmapsheet + "A11:F197" + "] ORDER BY 5";
                cmdExcel.CommandText = "SELECT * From [" + Roadmapsheet + "]";
                oda.SelectCommand = cmdExcel;
                oda.Fill(dtRoadmap);

                cmdExcel.CommandText = "SELECT * From [" + Roadmapcolor + "]";
                oda.SelectCommand = cmdExcel;
                oda.Fill(dtColor);

                cmdExcel.CommandText = "SELECT * From [" + Roadmapresource + "]";
                oda.SelectCommand = cmdExcel;
                oda.Fill(dtResource);


                connExcel.Close();
                var directory = Path.Combine(HttpContext.Current.Server.MapPath("~/CSV/CurrentFile"));

                DirectoryInfo d = new DirectoryInfo(directory);//Assuming Test is your Folder

                FileInfo[] file = d.GetFiles("Global_IT_Roadmap.xlsx");
                FileInfo excelfile = file[0];
                ExcelPackage xlPackage = new ExcelPackage(excelfile, false);

                ExcelWorksheet objSht = xlPackage.Workbook.Worksheets["Color"];
                int maxRow = dtColor.Rows.Count + 1;
                int maxCol = 6;


                OfficeOpenXml.ExcelRange range = objSht.Cells[1, 1, maxRow, maxCol];


                for (int i = 2; i <= maxRow; i++)
                {

                    string color = range[i, 2].Style.Fill.BackgroundColor.Rgb;
                    excelColordata = new ExcelColordata();
                    excelColordata.program_name = Convert.ToString(range[i, 1].Value).Trim();
                    excelColordata.project_color = range[i, 2].Style.Fill.BackgroundColor.Rgb;

                    excelColordata.region_name = Convert.ToString(range[i, 3].Value).Trim();
                    excelColordata.region_color = range[i, 4].Style.Fill.BackgroundColor.Rgb;

                    excelColordata.resource_name = Convert.ToString(range[i, 5].Value).Trim();
                    excelColordata.resource_color = range[i, 6].Style.Fill.BackgroundColor.Rgb;
                    lstExcelColordata.Add(excelColordata);

                }

                foreach (DataRow dr in dtResource.Rows)
                {
                    excelResourcedata = new ExcelResourcedata();
                    excelResourcedata.program_name = Convert.ToString(dr[0]).Trim();
                    excelResourcedata.resource_name = Convert.ToString(dr[3]).Trim();
                    lstExcelResourcedata.Add(excelResourcedata);
                    if (Convert.ToString(dr[6]) != "")
                    {
                        excelResourcedata = new ExcelResourcedata();
                        excelResourcedata.program_name = Convert.ToString(dr[0]).Trim();
                        excelResourcedata.resource_name = Convert.ToString(dr[6]).Trim();
                        lstExcelResourcedata.Add(excelResourcedata);
                    }
                }

                for (int i = 0; i < 9; i++)
                {
                    DataRow row = dtRoadmap.Rows[0];
                    dtRoadmap.Rows.Remove(row);
                }

                foreach (DataRow dr in dtRoadmap.Rows)
                {
                    if (dr[0].ToString() == "")
                    {
                        break;
                    }
                    else
                    {
                        var resourcelist = lstExcelResourcedata.Where(x => x.program_name == Convert.ToString(dr[0]).Trim())
                              .ToList();
                        for (int i = 0; i < resourcelist.Count; i++)
                        {
                            excelRoadMapdata = new ExcelRoadMapdata();
                            excelRoadMapdata.program_name = Convert.ToString(dr[0]).Trim();
                            excelRoadMapdata.region_name = Convert.ToString(dr[2]).Trim(); ;
                            excelRoadMapdata.country_name = Convert.ToString(dr[3]).Trim();
                            excelRoadMapdata.start_date = (Convert.ToDateTime(Convert.ToString(dr[4]).Trim().Replace('/', '-'))).ToString("dd-MM-yyyy");
                            excelRoadMapdata.end_date = (Convert.ToDateTime(Convert.ToString(dr[5]).Trim().Replace('/', '-'))).ToString("dd-MM-yyyy");                            //excelRoadMapdata.start_date = Convert.ToString(dr[4]).Trim().Replace('/', '-');
                           // excelRoadMapdata.start_date = Convert.ToString(dr[4]).Trim().Replace('/', '-');

                            //  excelRoadMapdata.end_date = Convert.ToString(dr[5]).Trim().Replace('/', '-');
                            excelRoadMapdata.resource_name = Convert.ToString(resourcelist[i].resource_name.Trim()); ;
                            lstExcelRoadMapdata.Add(excelRoadMapdata);
                        }

                    }
                }

                excelData = new ExcelData();
                lstExcelRoadMapdata = lstExcelRoadMapdata.OrderBy(o => Convert.ToDateTime(o.start_date)).ToList();
                excelData.excelRoadMapdata = lstExcelRoadMapdata;
                excelData.excelColordata = lstExcelColordata;

                lstExcelData.Add(excelData);
                //excelData = new ExcelData();                
                //  lstExcelData.Add(excelData);              

            }
            catch (Exception ex)
            {
                excelData = new ExcelData();
                excelData.excelColordata = lstExcelColordata;

                excelRoadMapdata = new ExcelRoadMapdata();
                excelRoadMapdata.program_name = ex.ToString();
                lstExcelRoadMapdata.Add(excelRoadMapdata);

                excelData.excelRoadMapdata = lstExcelRoadMapdata;

                lstExcelData.Add(excelData);
                connExcel.Close();
            }
            return lstExcelData;

        }
        
        public  void uploadSharePointFileFunction()
        {
            string result = "success";
            try
            {
                string Username = Decryption.DecryptNew(System.Configuration.ConfigurationManager.AppSettings["UserName"].ToString());
                string Password = Decryption.DecryptNew(System.Configuration.ConfigurationManager.AppSettings["Password"].ToString());

                var destpath = Path.Combine(HttpContext.Current.Server.MapPath("~/CSV/CurrentFile"), "Global_IT_Roadmap.xlsx");

                WebClient webClient = new WebClient();
                webClient.Credentials = new NetworkCredential(Username, Password);
                webClient.OpenRead(SharepointPath);
                webClient.DownloadFile(SharepointPath, destpath);

               // var archivefileName = "Global_IT_Roadmap" + "_" + DateTime.Now.ToString("MMddyyhhmm");
                //var archivepath = Path.Combine(HttpContext.Current.Server.MapPath("~/CSV/Archive"), archivefileName + ".xlsx");

              //  webClient.DownloadFile(SharepointPath, archivepath);
                //return result;
            }
            catch (Exception ex)
            {
                writeLog(ex);
                result = ex.ToString();
               // return result;
            }

        }


        //get username,password and sharepoint path from web.config

     //<add key="SharepointPath" value="http://35.188.173.90/sharpoint/Global_IT_Roadmap.xlsx" />
     //<add key="UserName" value="EuhDnBgBLwVmPXycbY7dMnflEPRgXde5SC4Dpntlu6Y=" />
     //<add key="Password" value="sXtnRDRDdHH5SG8zWvdcDN0W+SCpws9odWzbW6QAPdY=" 
       
        [HttpPost]
        [ActionName("uploadSharePointFile")]
        public string uploadSharePointFile()
        {
            string result = "success";
            try
            {
                string Username = Decryption.DecryptNew(System.Configuration.ConfigurationManager.AppSettings["UserName"].ToString());
                string Password = Decryption.DecryptNew(System.Configuration.ConfigurationManager.AppSettings["Password"].ToString());

                var destpath = Path.Combine(HttpContext.Current.Server.MapPath("~/CSV/CurrentFile"), "Global_IT_Roadmap.xlsx");
    
                WebClient webClient = new WebClient();
                webClient.Credentials = new NetworkCredential(Username,Password);
                webClient.OpenRead(SharepointPath);
                webClient.DownloadFile(SharepointPath, destpath);

                //var archivefileName = "Global_IT_Roadmap" + "_" + DateTime.Now.ToString("MMddyyhhmm");
                //var archivepath = Path.Combine(HttpContext.Current.Server.MapPath("~/CSV/Archive"), archivefileName + ".xlsx");

                //webClient.DownloadFile(SharepointPath, archivepath);
                return result;
            }
            catch(Exception ex)
            {              
                writeLog(ex);
                result = ex.ToString();
                return result;
            }

        }

        [HttpPost]
        [ActionName("uploadSharePointFileByNetworkMethod")]
        public string uploadSharePointFileByNetworkMethod()
        {
            string result = "success";
            try
            {
                string Username = Decryption.DecryptNew(System.Configuration.ConfigurationManager.AppSettings["UserName"].ToString());
                string Password = Decryption.DecryptNew(System.Configuration.ConfigurationManager.AppSettings["Password"].ToString());

                var destpath = Path.Combine(HttpContext.Current.Server.MapPath("~/CSV/CurrentFile"), "Global_IT_Roadmap.xlsx");

                 var networkPath = SharepointPath;
                 var credentials = new NetworkCredential(Username, Password);
                 string[] fileList;

                  //Try Method 1 to get files from sharepoint path
                 using (new NetworkConnection(networkPath, credentials))
                 {
                      fileList = Directory.GetFiles(networkPath);
                 }
                 //Try Method 2 to copy files from sharepoint path

                 using (new NetworkConnection(networkPath, credentials))
                 {
                     File.Copy(SharepointPath, destpath);
                 }

                
                return result;
            }
            catch (Exception ex)
            {
                writeLog(ex);
                result= ex.ToString();
                return result;
            }

        }

        public void writeLog(Exception ex)
        {
            string filePath = Path.Combine(HttpContext.Current.Server.MapPath("~/Log"), "Error.txt");
            using (StreamWriter writer = new StreamWriter(filePath, true))
            {
                writer.WriteLine("-----------------------------------------------------------------------------");
                writer.WriteLine("Date : " + DateTime.Now.ToString());
                writer.WriteLine();

                while (ex != null)
                {
                    writer.WriteLine(ex.GetType().FullName);
                    writer.WriteLine("Message : " + ex.Message);
                    writer.WriteLine("StackTrace : " + ex.StackTrace);

                    ex = ex.InnerException;
                }
            }
        }

        [HttpPost]
        [ActionName("uploadHistoryFile")]
        public string uploadHistoryFile(string filename)
        {
            string result = "success";
            try
            {

                var archivepath = Path.Combine(HttpContext.Current.Server.MapPath("~/CSV/Archive"));
                // string[] filePaths = Directory.GetFiles(path, filename);
                DirectoryInfo d = new DirectoryInfo(archivepath);//Assuming Test is your Folder

                FileInfo[] archiveFiles = d.GetFiles(filename); //Getting Text files

                var destpath = Path.Combine(HttpContext.Current.Server.MapPath("~/CSV/CurrentFile"), "Global_IT_Roadmap.xlsx");
                FileInfo[] deleteFiles = d.GetFiles("Global_IT_Roadmap.xlsx");
                foreach (FileInfo file in archiveFiles)
                {
                    file.CreationTime = DateTime.Now;
                    file.CopyTo(destpath, true);
                }
                return result;
            }
            catch(Exception ex)
            {
                return ex.ToString();
            }
        }

        [HttpGet]
        [ActionName("getArchive")]
        public List<HistoryFiledata> GetAllHistory()
        {
            List<HistoryFiledata> lsthistoryFiledata = new List<HistoryFiledata>();
            HistoryFiledata historyFiledata;
            try
            {
                //  var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(HttpContext.Current.Server.MapPath("~/CSV/Archive"));
                //string[] filePaths = Directory.GetFiles(path, "*.xlsx");
                DirectoryInfo d = new DirectoryInfo(path);//Assuming Test is your Folder

                FileInfo[] Files = d.GetFiles("*.xlsx").OrderByDescending(p => p.CreationTime).ToArray(); //Getting Text files

                foreach (FileInfo file in Files)
                {
                    historyFiledata = new HistoryFiledata();
                    historyFiledata.file_name = file.Name;
                    lsthistoryFiledata.Add(historyFiledata);
                }

            }
            catch (Exception ex)
            {
                //OTIS_Subscription_API.App_Code.LoggerHelper.ExcpLogger("FormatController", "GetAllFormat", ex);
                lsthistoryFiledata = null;
            }
            return lsthistoryFiledata;

        }

        [NonAction]
        public string PostUpload()
        {
            var file = HttpContext.Current.Request.Files.Count > 0 ?
            HttpContext.Current.Request.Files[0] : null;
            if (file.ContentLength > 0)
            {
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(HttpContext.Current.Server.MapPath("~/CSV"), "sampleCSV.csv");
                file.SaveAs(path);
            }
            return "~/CSV" + file.FileName;
        }

        [HttpPost]
        [ActionName("uploadExcelFile")]
        public string uploadExcelFile()
        {
            var file = HttpContext.Current.Request.Files.Count > 0 ?
            HttpContext.Current.Request.Files[0] : null;
            var fileName = "";
            if (file.ContentLength > 0)
            {
                var splitfileextension = file.FileName.Split('.');
                fileName = Path.GetFileName(splitfileextension[0].Replace(' ', '_')) + "_" + DateTime.Now.ToString("MMddyyyyhhmmss") + "." + splitfileextension[1];
                var archivepath = Path.Combine(HttpContext.Current.Server.MapPath("~/CSV/Archive"), fileName);

                file.SaveAs(archivepath);
                var currentpath = Path.Combine(HttpContext.Current.Server.MapPath("~/CSV/CurrentFile"), "Global_IT_Roadmap.xlsx");
                file.SaveAs(currentpath);
            }
            return "~/CSV" + file.FileName;
        }

    }
}
