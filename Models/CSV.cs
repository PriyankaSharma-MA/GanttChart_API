using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GanttChart.Models
{
    public class CSVdata
    {
        public string project_name { get; set; }
        public string country_name { get; set; }
        public string region_name { get; set; }
        public string resource_name { get; set; }

        public string start_date { get; set; }
        public string end_date { get; set; }
    }
    public class ExcelRoadMapdata
    {
        public string project_name { get; set; }
        public string country_name { get; set; }
        public string region_name { get; set; }
        public string resource1_name { get; set; }
        public string resource2_name { get; set; }

        public string start_date { get; set; }
        public string end_date { get; set; }
    }
    public class ExcelColordata
    {
        public string project_name { get; set; }
        public string region_name { get; set; }
        public string resource_name { get; set; }

        public string project_color { get; set; }
        public string region_color { get; set; }
        public string resource_color { get; set; }
    }
    public class ExcelData
    {
        public  List<ExcelRoadMapdata> excelRoadMapdata { get; set; }
        public List<ExcelColordata> excelColordata { get; set; } 
    }
    public class HistoryFiledata
    {
        public string file_name { get; set; }
    }
    public class decrptdata
    {
      
    }
}