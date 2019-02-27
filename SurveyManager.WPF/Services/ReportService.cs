using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SurveyManager.WPF.Models;
using Microsoft.VisualBasic.FileIO;
using Word = Microsoft.Office.Interop.Word;

namespace SurveyManager.WPF.Services
{
    public class ReportService
    {
        private readonly string surveyName;
        private readonly string surveyDataLocation;
        private readonly string reportDataLocation;
        private readonly string reportsDestination;
        private object reportTemplate;
        private IList<SurveyEntry> surveyEntries;
        private IList<IndividualReport> individualReports;
        private IndividualReport currentIndividualReport;
        private int numberOfQuantitativeQuestions;
        private Word.Application app;
        private Word.Document doc;

        public ReportService(string surveyName, string surveyDataLocation, string reportDataLocation, string reportsDestination, string reportTemplate)
        {
            this.surveyName = surveyName;
            this.surveyDataLocation = surveyDataLocation;
            this.reportDataLocation = reportDataLocation;
            this.reportsDestination = reportsDestination;
            this.reportTemplate = reportTemplate;
        }

        public void GenerateIndividualReport()
        {
            ReadSurveyData();
            ReadReportData();
            WriteIndividualReports();
            WriteStudentSurveyOnLecturerSummaryReport();
        }

        private void ReadSurveyData()
        {
            surveyEntries = new List<SurveyEntry>();

            using (var parser = new TextFieldParser(surveyDataLocation))
            {
                parser.Delimiters = new string[] { "," };
                parser.HasFieldsEnclosedInQuotes = true;

                string[] headers = parser.ReadFields();

                while (!parser.EndOfData)
                {
                    string[] row = parser.ReadFields();

                    switch (surveyName)
                    {
                        case SurveyName.StudentSurveyOnLecturer:
                            surveyEntries.Add(new StudentSurveyOnLecturerSurveyEntry(row));
                            break;
                        case SurveyName.UnitAndLecturerSurvey:
                            surveyEntries.Add(new UnitAndLecturerSurveyEntry(row));
                            break;
                    }
                }
            }
        }

        private void ReadReportData()
        {
            individualReports = new List<IndividualReport>();

            using (TextFieldParser parser = new TextFieldParser(reportDataLocation))
            {
                parser.Delimiters = new string[] { "," };
                parser.HasFieldsEnclosedInQuotes = true;

                string[] headers = parser.ReadFields();

                while (!parser.EndOfData)
                {
                    string[] columns = parser.ReadFields();

                    individualReports.Add(new IndividualReport(columns));
                }
            }
        }

        private void WriteIndividualReports()
        {
            app = new Word.Application();

            numberOfQuantitativeQuestions = surveyEntries.First().Questions.Where(q => q.IsQuantitative == true).Count();

            foreach (var individualReport in individualReports)
            {
                currentIndividualReport = individualReport;
                doc = app.Documents.Add(ref reportTemplate);

                individualReport.SurveyEntries = surveyEntries.Where(se => se.SurveyId == currentIndividualReport.ReportId);

                WriteBasicReportInformation();
                WriteSurveySpecificReportInformation();

                object filename = Path.Combine(reportsDestination, $"{currentIndividualReport.UnitCode} - {currentIndividualReport.Lecturer}.pdf");
                doc.SaveAs2(filename, Word.WdSaveFormat.wdFormatPDF);
                doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            }

            app.Quit();
        }

        private void WriteSurveySpecificReportInformation()
        {
            switch (surveyName)
            {
                case SurveyName.StudentSurveyOnLecturer:
                    WriteStudentSurveyOnLecturerInformation();
                    break;
                case SurveyName.UnitAndLecturerSurvey:
                    WriteUnitAndLecturerSurveyInformation();
                    break;
            }
        }

        private void WriteStudentSurveyOnLecturerInformation()
        {
            for (int questionNumber = 1; questionNumber <= 10; questionNumber++)
            {
                var numberOfStronglyAgree = currentIndividualReport.AllQuestions.Where(q => q.QuestionNumber == questionNumber && q.Answer == QuantitativeChoices.StronglyAgree).Count();
                var numberOfAgree = currentIndividualReport.AllQuestions.Where(q => q.QuestionNumber == questionNumber && q.Answer == QuantitativeChoices.Agree).Count();
                var numberOfNeutral = currentIndividualReport.AllQuestions.Where(q => q.QuestionNumber == questionNumber && q.Answer == QuantitativeChoices.Neutral).Count();
                var numberOfDisagree = currentIndividualReport.AllQuestions.Where(q => q.QuestionNumber == questionNumber && q.Answer == QuantitativeChoices.Disagree).Count();
                var numberOfStronglyDisagree = currentIndividualReport.AllQuestions.Where(q => q.QuestionNumber == questionNumber && q.Answer == QuantitativeChoices.StronglyDisagree).Count();
                var numberOfSkipped = currentIndividualReport.AllQuestions.Where(q => q.QuestionNumber == questionNumber && q.Answer == QuantitativeChoices.Skipped).Count();

                ReplaceText($"[Q{questionNumber}SA]", $"{numberOfStronglyAgree.ToString()} ({GetPercentage((double)numberOfStronglyAgree / currentIndividualReport.ClassSize)})");
                ReplaceText($"[Q{questionNumber}A]", $"{numberOfAgree.ToString()} ({GetPercentage((double)numberOfAgree / currentIndividualReport.ClassSize)})");
                ReplaceText($"[Q{questionNumber}N]", $"{numberOfNeutral.ToString()} ({GetPercentage((double)numberOfNeutral / currentIndividualReport.ClassSize)})");
                ReplaceText($"[Q{questionNumber}D]", $"{numberOfDisagree.ToString()} ({GetPercentage((double)numberOfDisagree / currentIndividualReport.ClassSize)})");
                ReplaceText($"[Q{questionNumber}SD]", $"{numberOfStronglyDisagree.ToString()} ({GetPercentage((double)numberOfStronglyDisagree / currentIndividualReport.ClassSize)})");
                ReplaceText($"[Q{questionNumber}S]", $"{numberOfSkipped.ToString()} ({GetPercentage((double)numberOfSkipped / currentIndividualReport.ClassSize)})");
            }

            var strengths = currentIndividualReport.AllQuestions.Where(q => q.QuestionNumber == 11).Select(q => q.Answer).ToList();
            var suggestions = currentIndividualReport.AllQuestions.Where(q => q.QuestionNumber == 12).Select(q => q.Answer).ToList();

            ReplaceText("[Q11Answer]", ConvertListToLines(strengths));
            ReplaceText("[Q12Answer]", ConvertListToLines(suggestions));

            ReplaceText("[SA]", GetPercentage((double)currentIndividualReport.AllQuestions.Where(q => q.Answer == QuantitativeChoices.StronglyAgree).Count() / (currentIndividualReport.SurveyEntries.Count() * numberOfQuantitativeQuestions)));
            ReplaceText("[A]", GetPercentage((double)currentIndividualReport.AllQuestions.Where(q => q.Answer == QuantitativeChoices.Agree).Count() / (currentIndividualReport.SurveyEntries.Count() * numberOfQuantitativeQuestions)));
            ReplaceText("[N]", GetPercentage((double)currentIndividualReport.AllQuestions.Where(q => q.Answer == QuantitativeChoices.Neutral).Count() / (currentIndividualReport.SurveyEntries.Count() * numberOfQuantitativeQuestions)));
            ReplaceText("[D]", GetPercentage((double)currentIndividualReport.AllQuestions.Where(q => q.Answer == QuantitativeChoices.Disagree).Count() / (currentIndividualReport.SurveyEntries.Count() * numberOfQuantitativeQuestions)));
            ReplaceText("[SD]", GetPercentage((double)currentIndividualReport.AllQuestions.Where(q => q.Answer == QuantitativeChoices.StronglyDisagree).Count() / (currentIndividualReport.SurveyEntries.Count() * numberOfQuantitativeQuestions)));
            ReplaceText("[S]", GetPercentage((double)currentIndividualReport.AllQuestions.Where(q => q.Answer == QuantitativeChoices.Skipped).Count() / (currentIndividualReport.SurveyEntries.Count() * numberOfQuantitativeQuestions)));

            currentIndividualReport.TotalPerformance = (double)currentIndividualReport.AllQuestions.Where(q => q.Answer == QuantitativeChoices.StronglyAgree || q.Answer == QuantitativeChoices.Agree).Count() / (currentIndividualReport.SurveyEntries.Count() * numberOfQuantitativeQuestions);
            ReplaceText("[Score]", GetPercentage(currentIndividualReport.TotalPerformance));
        }

        private void WriteUnitAndLecturerSurveyInformation()
        {

        }

        private void WriteBasicReportInformation()
        {
            ReplaceText("[StudyTerm]", currentIndividualReport.StudyTerm);
            ReplaceText("[LecturerName]", currentIndividualReport.Lecturer);
            ReplaceText("[Unit]", $"{currentIndividualReport.UnitCode.ToString()} - {currentIndividualReport.UnitName}");
            ReplaceText("[Population]", currentIndividualReport.ClassSize.ToString());
            ReplaceText("[Response]", $"{currentIndividualReport.SurveyEntries.Count().ToString()} ({GetPercentage((double)currentIndividualReport.SurveyEntries.Count() / currentIndividualReport.ClassSize)})");

            SetFlags();

            if (currentIndividualReport.Flags.Count == 0)
                ReplaceText("[Flag]", "No flag.");
            else
                ReplaceText("[Flag]", ConvertListToLines(currentIndividualReport.Flags));
        }

        private void WriteStudentSurveyOnLecturerSummaryReport()
        {
            var stringBuilder = new StringBuilder();

            IEnumerable<string> header = new List<string>()
            {
                "Unit Code",
                "Unit Name",
                "Lecturer",
                "Overall Performance Indicator (%)",
                "Status",
                "Number of Flag",
                "Response Rate (%)",
                "Response",
                "Population",
                "Question 1",
                "Question 2",
                "Question 3",
                "Question 4",
                "Question 5",
                "Question 6",
                "Question 7",
                "Question 8",
                "Question 9",
                "Question 10"
            };
            stringBuilder.AppendLine(string.Join(",", header.ToArray()));

            IList<string> body = new List<string>();
            foreach (var individualReport in individualReports)
            {
                currentIndividualReport = individualReport;

                var totalPerformance = Math.Round(100 * currentIndividualReport.TotalPerformance, 0);

                string status;
                if (totalPerformance < 80)
                    if (totalPerformance < 70)
                        status = "poor";
                    else
                        status = "alert";
                else
                    status = "good";

                body.Clear();
                body.Add(individualReport.UnitCode);
                body.Add(individualReport.UnitName);
                body.Add(individualReport.Lecturer);
                body.Add(totalPerformance.ToString());
                body.Add(status);
                body.Add(individualReport.Flags.Count.ToString());
                body.Add(GetPercentage((double)currentIndividualReport.SurveyEntries.Count() / currentIndividualReport.ClassSize));
                body.Add(currentIndividualReport.SurveyEntries.Count().ToString());
                body.Add(currentIndividualReport.ClassSize.ToString());

                for (int questionNumber = 1; questionNumber <= 10; questionNumber++)
                {
                    var performance = currentIndividualReport.AllQuestions.Where(q => q.QuestionNumber == questionNumber && (q.Answer == QuantitativeChoices.StronglyAgree || q.Answer == QuantitativeChoices.Agree)).Count();
                    body.Add(performance.ToString());
                }
                stringBuilder.AppendLine(string.Join(",", body.ToArray()));
            }

            File.WriteAllText(Path.Combine(reportsDestination, "SummaryReport.csv"), stringBuilder.ToString());
        }

        private void ReplaceText(object placeholder, string value)
        {
            var range = doc.Content;
            range.Find.MatchWholeWord = true;
            range.Find.MatchCase = true;

            if (!range.Find.Execute(placeholder))
                return;

            range.Delete();
            range.Text = value ?? string.Empty;
        }

        private string GetPercentage(double value)
        {
            return Math.Round(value * 100, 0).ToString() + "%";
        }

        private string ConvertListToLines(IList<string> values)
        {
            var validAnswer = values.Where(v => !string.IsNullOrEmpty(v)).ToArray();

            if (validAnswer.Count() == 0)
                return string.Empty;

            return string.Join("\n", validAnswer);
        }

        private void SetFlags()
        {
            if (currentIndividualReport.SurveyEntries.Count() == 0)
                currentIndividualReport.Flags.Add("No response collected");

            if (currentIndividualReport.ClassSize < 10)
                currentIndividualReport.Flags.Add("The class size is less than 10 students.");

            if (((double)currentIndividualReport.SurveyEntries.Count() / currentIndividualReport.ClassSize) < 0.2)
                currentIndividualReport.Flags.Add("The response rate is less than 20%.");
        }
    }
}
