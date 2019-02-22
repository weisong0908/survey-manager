using SurveyManager.WPF.Models;
using SurveyManager.WPF.Views;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace SurveyManager.WPF.Services
{
    public class WindowService
    {
        public void ShowWindow(string surveyName)
        {
            var window = new IndividualSurveyWindow(surveyName);

            window.ShowDialog();
        }
    }
}
