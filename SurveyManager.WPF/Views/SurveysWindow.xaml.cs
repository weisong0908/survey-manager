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
    /// Interaction logic for SurveysWindow.xaml
    /// </summary>
    public partial class SurveysWindow : Window
    {
        public SurveysWindowViewModel ViewModel { get { return DataContext as SurveysWindowViewModel; } set { DataContext = value; } }

        public SurveysWindow()
        {
            InitializeComponent();

            ViewModel = new SurveysWindowViewModel();
        }

        private void GoToStudentSurveyOnLecturerWindow(object sender, RoutedEventArgs e)
        {
            ViewModel.GoToStudentSurveyOnLecturerWindow();
        }

        private void GoToUnitAndLecturerSurveyWindow(object sender, RoutedEventArgs e)
        {
            ViewModel.GoToUnitAndLecturerSurveyWindow();
        }
    }
}
