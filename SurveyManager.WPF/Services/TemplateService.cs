﻿using SurveyManager.WPF.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SurveyManager.WPF.Services
{
    public class TemplateService
    {
        private readonly string templateParentFolder;
        private readonly string surveyDataTemplateFileName = "SurveyDataTemplate.csv";
        private readonly string reportDataTemplateFileName = "ReportDataTemplate.csv";
        private readonly string ReportTemplateFileName = "ReportTemplate.dotx";

        public string SurveyDataTemplate { get; private set; }
        public string ReportDataTemplate { get; private set; }
        public string ReportTemplate { get; private set; }

        public TemplateService(string surveyName)
        {
            switch (surveyName)
            {
                case SurveyName.StudentSurveyOnLecturer:
                    templateParentFolder = @"\\csing.navitas.local\shared\Documents\Quality Assurance\Survey\Templates\CUStudentSurveyOnLecturer";
                    break;
                case SurveyName.UnitAndLecturerSurvey:
                    templateParentFolder = @"\\csing.navitas.local\shared\Documents\Quality Assurance\Survey\Templates\CCUnitAndLecturerSurvey";
                    break;
            }

            SurveyDataTemplate = Path.Combine(templateParentFolder, surveyDataTemplateFileName);
            ReportDataTemplate = Path.Combine(templateParentFolder, reportDataTemplateFileName);
            ReportTemplate = Path.Combine(templateParentFolder, ReportTemplateFileName);
        }

        public void ExportTemplates()
        {
            var folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.Description = "Export templates to folder";

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                File.Copy(SurveyDataTemplate, Path.Combine(folderBrowserDialog.SelectedPath, surveyDataTemplateFileName), overwrite: true);
                File.Copy(ReportDataTemplate, Path.Combine(folderBrowserDialog.SelectedPath, reportDataTemplateFileName), overwrite: true);
            }
        }
    }
}
