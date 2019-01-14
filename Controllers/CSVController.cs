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
//using  Microsoft.Office.Interop.Excel;
//using ExcelLibrary;

//using CSharpJExcel.Jxl;
using OfficeOpenXml;



namespace GanttChart.Controllers
{
    public class CSVController : ApiController
    {
        string CSVPath = System.Configuration.ConfigurationManager.AppSettings["CSVPath"];
        string SharepointPath = System.Configuration.ConfigurationManager.AppSettings["SharepointPath"];
        string Sharepointfilename = System.Configuration.ConfigurationManager.AppSettings["Sharepointfilename"]; 

        [NonAction]
        public List<CSVdata> ReadCsvFile()
        {

            CSVdata csvdata;
            List<CSVdata> csvdataList = new List<CSVdata>();

            using (var reader = new StreamReader(Path.Combine(HttpContext.Current.Server.MapPath("~/CSV/CurrentFile"), "sampleCSV.csv")))
            {

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');


                    csvdata = new CSVdata();
                    csvdata.program_name = Convert.ToString(values[0]).Trim();
                    csvdata.region_name = Convert.ToString(values[1]).Trim();
                    csvdata.country_name = Convert.ToString(values[2]).Trim();
                    csvdata.resource_name = Convert.ToString(values[3]).Trim();


                    csvdata.start_date = Convert.ToString(values[4].Replace('/', '-')).Trim();
                    csvdata.end_date = Convert.ToString(values[5].Replace('/', '-')).Trim();

                    csvdataList.Add(csvdata);
                }
            }

            return csvdataList.Skip(1).ToList();
        }
        [HttpGet]
        [ActionName("GetAllCSVData")]
        public List<CSVdata> GetAllCSVData([FromUri] string filename)
        {
            List<CSVdata> csvdataList = new List<CSVdata>();

            try
            {
                string requestPath = filename;
                csvdataList = ReadCsvFile();
            }
            catch (Exception ex)
            {
                //OTIS_Subscription_API.App_Code.LoggerHelper.ExcpLogger("FormatController", "GetAllFormat", ex);
                csvdataList = null;
            }
            return csvdataList;

        }
        [HttpPost]
        [ActionName("uploadSharePointFile")]
        public void uploadSharePointFile()
        {


            var destpath = Path.Combine(HttpContext.Current.Server.MapPath("~/CSV/CurrentFile"), "Global_IT_Roadmap.xlsx");

            WebClient webClient = new WebClient();
            webClient.DownloadFile(SharepointPath, destpath);
          
            var archivefileName = "Global_IT_Roadmap" + "_" + DateTime.Now.ToString("MMddyyhhmm");
            var archivepath = Path.Combine(HttpContext.Current.Server.MapPath("~/CSV/Archive"), archivefileName + ".xlsx");

            webClient.DownloadFile(SharepointPath, archivepath);
        }

        [HttpPost]
        [ActionName("uploadHistoryFile")]
        public void uploadHistoryFile(string filename)
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
        }
        [HttpGet]
        [ActionName("GetArchive")]
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
        [HttpGet]
        [ActionName("GetAllExcelData")]
        public List<ExcelData> GetAllExcelData([FromUri] string filename)
        {

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

                string conStr = "", Extension = ".xlsx";
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
                cmdExcel.CommandText = "SELECT * From [" + Roadmapsheet  + "]";
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
                    if (Convert.ToString(dr[5]) != "")
                    {
                        excelResourcedata = new ExcelResourcedata();
                        excelResourcedata.program_name = Convert.ToString(dr[0]).Trim();
                        excelResourcedata.resource_name = Convert.ToString(dr[5]).Trim();
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
                            excelRoadMapdata.start_date = (Convert.ToDateTime(Convert.ToString(dr[4]).Trim())).ToString("dd-MM-yyyy");
                            excelRoadMapdata.end_date = (Convert.ToDateTime(Convert.ToString(dr[5]).Trim())).ToString("dd-MM-yyyy");
                            excelRoadMapdata.resource_name = Convert.ToString(resourcelist[i].resource_name.Trim()); ;
                            lstExcelRoadMapdata.Add(excelRoadMapdata);
                        }

                    }
                }

                excelData = new ExcelData();
                excelData.excelRoadMapdata = lstExcelRoadMapdata;
                excelData.excelColordata = lstExcelColordata;

                lstExcelData.Add(excelData);
                //excelData = new ExcelData();                
                //  lstExcelData.Add(excelData);              

            }
            catch (Exception ex)
            {

            }
            return lstExcelData;

        }

        [NonAction]
        // [ActionName("UploadFile")]
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
        [ActionName("UploadExcelFile")]
        public string PostUploadExcel()
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
