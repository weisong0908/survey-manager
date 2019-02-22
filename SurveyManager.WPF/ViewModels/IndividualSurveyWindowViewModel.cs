﻿using SurveyManager.WPF.Services;
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

        public IndividualSurveyWindowViewModel(string surveyName)
        {
            fileService = new FileService(surveyName);

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
    }
}
