using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GanttChart.Models
{
    public class CSVdata
    {
        public string program_name { get; set; }
        public string country_name { get; set; }
        public string region_name { get; set; }
        public string resource_name { get; set; }

        public string start_date { get; set; }
        public string end_date { get; set; }
    }
    public class ExcelRoadMapdata
    {
        public string program_name { get; set; }
        public string country_name { get; set; }
        public string region_name { get; set; }
    
        public string start_date { get; set; }
        public string end_date { get; set; }

        public string resource_name { get; set; }
        //public string resource1_name { get; set; }
       // public string resource2_name { get; set; }
    }
    public class ExcelColordata
    {
        public string program_name { get; set; }
        public string region_name { get; set; }
        public string resource_name { get; set; }

        public string project_color { get; set; }
        public string region_color { get; set; }
        public string resource_color { get; set; }
    }
    public class ExcelResourcedata
    {
        public string program_name { get; set; }
        public string resource_name { get; set; }
      //  public string resource1_name { get; set; }
       // public string resource2_name { get; set; }
    }
    public class ExcelData
    {
        public  List<ExcelRoadMapdata> excelRoadMapdata { get; set; }
        public List<ExcelColordata> excelColordata { get; set; }
        public List<ExcelResourcedata> excelResourcedata { get; set; } 
    }
    public class HistoryFiledata
    {
        public string file_name { get; set; }
    }
    public class programlist
    {
        public string program_id { get; set; }
        public string program_name { get; set; }
    }
}