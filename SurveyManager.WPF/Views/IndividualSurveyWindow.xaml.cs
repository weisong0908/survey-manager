using SurveyManager.WPF.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace SurveyManager.WPF.Views
{
    /// <summary>
    /// Interaction logic for IndividualSurveyWindow.xaml
    /// </summary>
    public partial class IndividualSurveyWindow : Window
    {
        public IndividualSurveyWindowViewModel ViewModel { get { return DataContext as IndividualSurveyWindowViewModel; } set { DataContext = value; } }

        public IndividualSurveyWindow(string surveyName)
        {
            InitializeComponent();

            ViewModel = new IndividualSurveyWindowViewModel(surveyName);
        }

        private void ExportTemplate(object sender, RoutedEventArgs e)
        {
            ViewModel.ExportTemplates();
        }

        private void ImportSurveyData(object sender, RoutedEventArgs e)
        {
            ViewModel.ImportSurveyData();
        }

        private void ImportReportData(object sender, RoutedEventArgs e)
        {
            ViewModel.ImportReportData();
        }

        private void GenerateReport(object sender, RoutedEventArgs e)
        {

        }
    }
}
