using SurveyManager.WPF.Models;
using SurveyManager.WPF.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SurveyManager.WPF.ViewModels
{
    public class SurveysWindowViewModel : BaseViewModel
    {
        private readonly WindowService windowService;

        public SurveysWindowViewModel()
        {
            windowService = new WindowService();
        }

        public void GoToStudentSurveyOnLecturerWindow()
        {
            windowService.ShowWindow(SurveyName.StudentSurveyOnLecturer);
        }

        public void GoToUnitAndLecturerSurveyWindow()
        {
            windowService.ShowWindow(SurveyName.UnitAndLecturerSurvey);
        }
    }
}
