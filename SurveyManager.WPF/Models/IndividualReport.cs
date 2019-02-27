using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SurveyManager.WPF.Models
{
    public class IndividualReport
    {
        public int ReportId { get; set; }
        public string UnitCode { get; set; }
        public string UnitName { get; set; }
        public string Lecturer { get; set; }
        public string StudyTerm { get; set; }
        public int ClassSize { get; set; }
        public int Response { get; set; }
        public IEnumerable<SurveyEntry> SurveyEntries { get; set; }
        public IEnumerable<SurveyQuestion> AllQuestions { get { return SurveyEntries.SelectMany(se => se.Questions); } }
        public double TotalPerformance { get; set; }
        public IList<string> Flags { get; set; }

        public IndividualReport(string[] columns)
        {
            SurveyEntries = new List<SurveyEntry>();
            Flags = new List<string>();

            ReportId = (int.TryParse(columns[0], out int id)) ? id : 0;
            UnitCode = columns[1];
            UnitName = columns[2];
            Lecturer = columns[3];
            ClassSize = (int.TryParse(columns[4], out int classSize)) ? classSize : 0;
            StudyTerm = columns[5];
        }
    }
}
