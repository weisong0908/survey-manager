using SurveyManager.WPF.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SurveyManager.WPF.ViewModels
{
    public class IndividualSurveyWindowViewModel : BaseViewModel
    {
        private readonly DataService dataService;

        public string SurveyName { get; private set; }

        private string _surveyDataLocation;
        public string SurveyDataLocation
        {
            get { return _surveyDataLocation; }
            set { SetValue(ref _surveyDataLocation, value); }
        }

        private string _reportDataLocation;
        public string ReportDataLocation
        {
            get { return _reportDataLocation; }
            set { SetValue(ref _reportDataLocation, value); }
        }

        private string _reportsLocation;
        public string ReportsLocation
        {
            get { return _reportsLocation; }
            set { SetValue(ref _reportsLocation, value); }
        }

        public IndividualSurveyWindowViewModel(string surveyName)
        {
            dataService = new DataService(surveyName);

            SurveyName = surveyName;
        }

        public void ExportTemplates()
        {
            dataService.ExportTemplates();
        }

        public void ImportSurveyData()
        {
            var location = dataService.ImportData(DataService.DataType.SurveyData);

            if (string.IsNullOrEmpty(location))
                return;
            SurveyDataLocation = location;
        }

        public void ImportReportData()
        {
            var location = dataService.ImportData(DataService.DataType.ReportData);

            if (string.IsNullOrEmpty(location))
                return;
            ReportDataLocation = location;
        }
    }
}
