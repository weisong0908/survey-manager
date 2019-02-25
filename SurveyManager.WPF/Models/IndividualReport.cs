using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SurveyManager.WPF.Models
{
    public class IndividualReport
    {
        public int Id { get; set; }
        public string UnitCode { get; set; }
        public string UnitName { get; set; }
        public string Lecturer { get; set; }
        public string StudyTerm { get; set; }
        public int ClassSize { get; set; }
        public int Response { get; set; }
        public IList<string> Alerts { get; set; }

        public IndividualReport()
        {
            Alerts = new List<string>();
        }
    }
}
