using SurveyManager.WPF.Models;
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
        private readonly FileService fileService;
        private WindowService windowService;
        private ReportService reportService;

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

        private string _reportsDestination;
        public string ReportsDestination
        {
            get { return _reportsDestination; }
            set { SetValue(ref _reportsDestination, value); }
        }

        private bool _isProgressBarRun;
        public bool IsProgressBarRun
        {
            get { return _isProgressBarRun; }
            set { SetValue(ref _isProgressBarRun, value); }
        }

        public IndividualSurveyWindowViewModel(string surveyName)
        {
            fileService = new FileService(surveyName);
            windowService = new WindowService();

            SurveyName = surveyName;
        }

        public void ExportTemplates()
        {
            fileService.ExportTemplates();
        }

        public void ImportSurveyData()
        {
            var location = fileService.ImportData(FileService.DataType.SurveyData);

            if (string.IsNullOrEmpty(location))
                return;
            SurveyDataLocation = location;
        }

        public void ImportReportData()
        {
            var location = fileService.ImportData(FileService.DataType.ReportData);

            if (string.IsNullOrEmpty(location))
                return;
            ReportDataLocation = location;
        }

        public void SetReportsDestination()
        {
            var location = fileService.SetReportsDestination();

            if (string.IsNullOrEmpty(location))
                return;
            ReportsDestination = location;
        }

        public async void GenerateReports()
        {
            string individualReportTemplateLocation;
            switch (SurveyName)
            {
                case SurveyNames.StudentSurveyOnLecturer:
                    individualReportTemplateLocation = IndividualReportTemplateLocations.StudentSurveyOnLecturer;
                    break;
                case SurveyNames.UnitAndLecturerSurvey:
                    individualReportTemplateLocation = IndividualReportTemplateLocations.UnitAndLecturerSurvey;
                    break;
                default:
                    individualReportTemplateLocation = string.Empty;
                    break;
            }

            reportService = new ReportService(SurveyName, _surveyDataLocation, _reportDataLocation, _reportsDestination, individualReportTemplateLocation);

            IsProgressBarRun = true;

            Task generateReportTask = new Task(() => reportService.GenerateIndividualReport());

            generateReportTask.Start();

            await generateReportTask;

            IsProgressBarRun = false;

            windowService.ShowMessageBox("The reports have been generated.", "Complete");
        }
    }
}
