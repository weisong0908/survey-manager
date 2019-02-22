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
        private readonly TemplateService templateService;

        public string SurveyName { get; private set; }

        public IndividualSurveyWindowViewModel(string surveyName)
        {
            templateService = new TemplateService(surveyName);

            SurveyName = surveyName;
        }

        public void ExportTemplates()
        {
            templateService.ExportTemplates();
        }
    }
}
