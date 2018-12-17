using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using GanttChart.Models;
using System.IO;
using System.Web;

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
           
            using (var reader = new StreamReader(Path.Combine(HttpContext.Current.Server.MapPath("~/CSV"), "sampleCSV.csv")))
            {

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');


                    csvdata = new CSVdata();
                    csvdata.project_name = Convert.ToString(values[0]).Trim();
                    csvdata.country_name = Convert.ToString(values[1]).Trim();
                    csvdata.resource_name = Convert.ToString(values[2]).Trim();


                    csvdata.start_date = Convert.ToString(values[3].Replace('/', '-')).Trim();
                    csvdata.end_date = Convert.ToString(values[4].Replace('/', '-')).Trim();

                    csvdataList.Add(csvdata);
                }
            }

            return csvdataList.Skip(1).ToList();
        }
        [HttpGet]
        public List<CSVdata> GetAllCSVData()
        {
            List<CSVdata> csvdataList = new List<CSVdata>();

            try
            {
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
        [ActionName("UploadFile")]
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



    }
}
