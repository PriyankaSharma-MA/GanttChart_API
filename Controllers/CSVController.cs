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
                    csvdata.project_name = Convert.ToString(values[0]).Trim();
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
        [ActionName("uploadHistoryFile")]
        public void uploadHistoryFile(string filename)
        {

            var archivepath = Path.Combine(HttpContext.Current.Server.MapPath("~/CSV/Archive"));
            // string[] filePaths = Directory.GetFiles(path, filename);
            DirectoryInfo d = new DirectoryInfo(archivepath);//Assuming Test is your Folder

            FileInfo[] archiveFiles = d.GetFiles(filename); //Getting Text files

            var destpath = Path.Combine(HttpContext.Current.Server.MapPath("~/CSV/CurrentFile"), "Roadmap.xlsx");
            FileInfo[] deleteFiles = d.GetFiles("Roadmap.xlsx");
            //foreach (FileInfo file in deleteFiles)
            //{
            //    file.Delete();

            //}
            foreach (FileInfo file in archiveFiles)
            {
                // System.IO.File.Copy(file.Name.Split('.')[0], path1.ToString());
                //file.MoveTo(destpath);


                file.CreationTime = DateTime.Now;
                file.CopyTo(destpath, true);
                // file.MoveTo(destpath);
                // file.Replace(destpath, "kk.xlsx");
                //    file.Replace(path1, path1);
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
  
            string Roadmapsheet = "Roadmap$";
            string Roadmapcolor = "Color$";

            ExcelRoadMapdata excelRoadMapdata;
            ExcelColordata excelColordata;
            ExcelData excelData;
            List<ExcelRoadMapdata> lstExcelRoadMapdata = new List<ExcelRoadMapdata>();
            List<ExcelColordata> lstExcelColordata = new List<ExcelColordata>();
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
                conStr = String.Format(conStr, Path.Combine(HttpContext.Current.Server.MapPath("~/CSV/CurrentFile"),filename ));
                OleDbConnection connExcel = new OleDbConnection(conStr);
                OleDbCommand cmdExcel = new OleDbCommand();
                OleDbDataAdapter oda = new OleDbDataAdapter();
                System.Data.DataTable dt = new System.Data.DataTable();
                System.Data.DataTable dtRoadmap = new System.Data.DataTable();
                System.Data.DataTable dtColor = new System.Data.DataTable();
                cmdExcel.Connection = connExcel;

                //Get the name of First Sheet
              

              
               

                //Read Data from First Sheet
                connExcel.Open();
                cmdExcel.CommandText = "SELECT * From [" + Roadmapsheet + "] ORDER BY 6";
                oda.SelectCommand = cmdExcel;
                oda.Fill(dtRoadmap);
             
                cmdExcel.CommandText = "SELECT * From [" + Roadmapcolor + "]";
                oda.SelectCommand = cmdExcel;
                oda.Fill(dtColor);
                connExcel.Close();
                 var directory = Path.Combine(HttpContext.Current.Server.MapPath("~/CSV/CurrentFile"));
    
                  DirectoryInfo d = new DirectoryInfo(directory);//Assuming Test is your Folder
             
                 FileInfo[] file = d.GetFiles("Roadmap.xlsx");
                 FileInfo excelfile = file[0];
                 ExcelPackage xlPackage = new ExcelPackage(excelfile,true);

                 ExcelWorksheet objSht = xlPackage.Workbook.Worksheets[2];
                 int maxRow = dtColor.Rows.Count+1;
                 int maxCol = 6;


                 OfficeOpenXml.ExcelRange range = objSht.Cells[1, 1, maxRow, maxCol];


                 for (int i = 2; i <= maxRow; i++)
                 {

                     string color = range[i, 2].Style.Fill.BackgroundColor.Rgb;
                     excelColordata = new ExcelColordata();
                     excelColordata.project_name = Convert.ToString(range[i, 1].Value).Trim();
                     excelColordata.project_color = range[i, 2].Style.Fill.BackgroundColor.Rgb;

                     excelColordata.region_name = Convert.ToString(range[i, 3].Value).Trim();
                     excelColordata.region_color = range[i, 4].Style.Fill.BackgroundColor.Rgb;

                     excelColordata.resource_name = Convert.ToString(range[i, 5].Value).Trim();
                     excelColordata.resource_color = range[i, 6].Style.Fill.BackgroundColor.Rgb;
                     lstExcelColordata.Add(excelColordata);

                 }
                 //foreach (DataRow dr in dtColor.Rows)
                 //{

                 //    excelColordata = new ExcelColordata();
                 //    excelColordata.project_name = Convert.ToString(dr[0]);
                 //    excelColordata.project_color = Convert.ToString(dr[1]);

                 //    excelColordata.region_name = Convert.ToString(dr[2]);
                 //    excelColordata.region_color = Convert.ToString(dr[3]);

                 //    excelColordata.resource_name = Convert.ToString(dr[4]);
                 //    excelColordata.resource_color = Convert.ToString(dr[5]);
                 //    lstExcelColordata.Add(excelColordata);
                 //}
                     foreach (DataRow dr in dtRoadmap.Rows)
                     {


                         excelRoadMapdata = new ExcelRoadMapdata();
                         excelRoadMapdata.project_name = Convert.ToString(dr[0]).Trim();
                         excelRoadMapdata.region_name = Convert.ToString(dr[1]).Trim(); ;
                         excelRoadMapdata.country_name = Convert.ToString(dr[2]).Trim(); 
                         excelRoadMapdata.resource1_name = Convert.ToString(dr[3]).Trim(); 
                         excelRoadMapdata.resource2_name = Convert.ToString(dr[4]).Trim(); 
                         excelRoadMapdata.start_date = (Convert.ToDateTime(Convert.ToString(dr[5]).Trim())).ToString("dd-MM-yyyy");
                         excelRoadMapdata.end_date = (Convert.ToDateTime(Convert.ToString(dr[6]).Trim())).ToString("dd-MM-yyyy");
                         lstExcelRoadMapdata.Add(excelRoadMapdata);
                         if (Convert.ToString(dr[4]) != "")
                         {
                             excelRoadMapdata = new ExcelRoadMapdata();
                             excelRoadMapdata.project_name = Convert.ToString(dr[0]).Trim();
                             excelRoadMapdata.region_name = Convert.ToString(dr[1]).Trim();
                             excelRoadMapdata.country_name = Convert.ToString(dr[2]).Trim();
                             excelRoadMapdata.resource1_name = Convert.ToString(dr[4]).Trim();
                             excelRoadMapdata.resource2_name = Convert.ToString(dr[4]).Trim();
                             excelRoadMapdata.start_date = (Convert.ToDateTime(Convert.ToString(dr[5]).Trim())).ToString("dd-MM-yyyy");
                             excelRoadMapdata.end_date = (Convert.ToDateTime(Convert.ToString(dr[6]).Trim())).ToString("dd-MM-yyyy");
                             lstExcelRoadMapdata.Add(excelRoadMapdata);
                         }
                     }
                     //lstExcelRoadMapdata.Sort();
              
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
                var path = Path.Combine(HttpContext.Current.Server.MapPath("~/CSV/Archive"), fileName);

                file.SaveAs(path);
                var path1 = Path.Combine(HttpContext.Current.Server.MapPath("~/CSV/CurrentFile"), "Roadmap.xlsx");
                file.SaveAs(path1);
            }
            return "~/CSV" + file.FileName;
        }
        public static StringBuilder Decryptpassword(string cipherText)
        {
            cipherText = Regex.Replace(cipherText, ".{6}", "$0,");
            string[] cahractergroup = cipherText.Split(',');
            StringBuilder decryptpassword = new StringBuilder();
            Dictionary<string, char> decryptcodelist = new Dictionary<string, char>();
            decryptcodelist.Add("uDFM45", 'a');
            decryptcodelist.Add("H21DGF", 'b');
            decryptcodelist.Add("FDH56D", 'c');
            decryptcodelist.Add("FGS546", 'd');
            decryptcodelist.Add("JUK4JH", 'e');
            decryptcodelist.Add("ERG54S", 'f');

            decryptcodelist.Add("T5H4FD", 'g');
            decryptcodelist.Add("RG641G", 'h');
            decryptcodelist.Add("RG4F4D", 'i');
            decryptcodelist.Add("RT56F6", 'j');
            decryptcodelist.Add("VCBC3B", 'k');
            decryptcodelist.Add("F8G9GF", 'l');

            decryptcodelist.Add("FD4CJS", 'm');
            decryptcodelist.Add("G423FG", 'n');
            decryptcodelist.Add("F45GC2", 'o');
            decryptcodelist.Add("TH5DF5", 'p');
            decryptcodelist.Add("CV4F6R", 'q');
            decryptcodelist.Add("XF64TS", 'r');

            decryptcodelist.Add("X78DGT", 's');
            decryptcodelist.Add("TH74SJ", 't');
            decryptcodelist.Add("bCX6DF", 'u');
            decryptcodelist.Add("FG65SD", 'v');
            decryptcodelist.Add("4KL45D", 'w');
            decryptcodelist.Add("GFH3F2", 'x');

            decryptcodelist.Add("GH56GF", 'y');
            decryptcodelist.Add("45T1FG", 'z');
            decryptcodelist.Add("D4G23D", '1');
            decryptcodelist.Add("GB56FG", '2');
            decryptcodelist.Add("sF45GF", '3');
            decryptcodelist.Add("P4FF12", '4');

            decryptcodelist.Add("F6DFG1", '5');
            decryptcodelist.Add("56FG4G", '6');
            decryptcodelist.Add("uSGFDG", '7');
            decryptcodelist.Add("FKHFDG", '8');
            decryptcodelist.Add("iFGJH6", '9');
            decryptcodelist.Add("87H8G7", '0');

            decryptcodelist.Add("G25GHF", '@');
            decryptcodelist.Add("45FGFH", '#');
            decryptcodelist.Add("75FG45", '$');
            decryptcodelist.Add("54GDH5", '*');
            decryptcodelist.Add("45F465", '(');
            decryptcodelist.Add("HG56FG", '.');

            decryptcodelist.Add("DF56H4", ',');
            decryptcodelist.Add("F5JHFH", '-');
            decryptcodelist.Add("sGF4HF", '=');
            decryptcodelist.Add("45GH45", '\\');
            decryptcodelist.Add("56H45G", '/');


            for (int i = 0; i < cahractergroup.Length - 1; i++)
            {
                decryptpassword = decryptpassword.Append(decryptcodelist[cahractergroup[i]]);


            }

            return decryptpassword;
        }

    }
}
