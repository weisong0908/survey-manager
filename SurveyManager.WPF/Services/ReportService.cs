using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SurveyManager.WPF.Services
{
    public class ReportService
    {
        private readonly string surveyDataLocation;
        private readonly string reportDataLocation;
        private readonly string reportsDestination;

        public ReportService(string surveyDataLocation, string reportDataLocation, string reportsDestination)
        {
            this.surveyDataLocation = surveyDataLocation;
            this.reportDataLocation = reportDataLocation;
            this.reportsDestination = reportsDestination;
        }
    }
}
