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
        public string resource_name { get; set; }

        public string start_date { get; set; }
        public string end_date { get; set; }
    }
}