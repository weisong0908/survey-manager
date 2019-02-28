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
        private object reportTemplateLocation;
        private IList<SurveyEntry> surveyEntries;
        private IList<IndividualReport> individualReports;
        private IndividualReport currentIndividualReport;
        private int numberOfQuantitativeQuestions;
        private Word.Application app;
        private Word.Document doc;
        public EventHandler<string> ProgressCompleted;

        public ReportService(string surveyName, string surveyDataLocation, string reportDataLocation, string reportsDestination, string reportTemplateLocation)
        {
            this.surveyName = surveyName;
            this.surveyDataLocation = surveyDataLocation;
            this.reportDataLocation = reportDataLocation;
            this.reportsDestination = reportsDestination;
            this.reportTemplateLocation = reportTemplateLocation;
        }

        public async Task GenerateIndividualReportAsync()
        {
            await Task.Run(() =>
                {
                    ReadSurveyData();
                    ReadReportData();
                    WriteIndividualReports();
                    WriteSummaryReport();
                });
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
                        case SurveyNames.StudentSurveyOnLecturer:
                            surveyEntries.Add(new StudentSurveyOnLecturerSurveyEntry(row));
                            break;
                        case SurveyNames.UnitAndLecturerSurvey:
                            surveyEntries.Add(new UnitAndLecturerSurveyEntry(row));
                            break;
                    }
                }
            }

            ProgressCompleted?.Invoke(this, "Survey data has been read");
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

            ProgressCompleted?.Invoke(this, "Report data has been read");
        }

        private void WriteIndividualReports()
        {
            app = new Word.Application();

            numberOfQuantitativeQuestions = surveyEntries.First().Questions.Where(q => q.IsQuantitative == true).Count();

            foreach (var individualReport in individualReports)
            {
                currentIndividualReport = individualReport;
                doc = app.Documents.Add(ref reportTemplateLocation);

                individualReport.SurveyEntries = surveyEntries.Where(se => se.SurveyId == currentIndividualReport.ReportId);

                WriteBasicReportInformation();
                WriteSurveySpecificReportInformation();

                object filename = Path.Combine(reportsDestination, $"{currentIndividualReport.UnitCode} - {currentIndividualReport.Lecturer}.pdf");
                doc.SaveAs2(filename, Word.WdSaveFormat.wdFormatPDF);
                doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);

                ProgressCompleted?.Invoke(this, $" Individual report for {currentIndividualReport.UnitCode} - {currentIndividualReport.Lecturer} has been completed");
            }

            app.Quit();
        }

        private void WriteSurveySpecificReportInformation()
        {
            switch (surveyName)
            {
                case SurveyNames.StudentSurveyOnLecturer:
                    WriteStudentSurveyOnLecturerInformation();
                    break;
                case SurveyNames.UnitAndLecturerSurvey:
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
            for (int questionNumber = 1; questionNumber <= 6; questionNumber++)
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

            ReplaceText("[Q7Answer]", ConvertListToLines(currentIndividualReport.AllQuestions.Where(q => q.QuestionNumber == 7).Select(q => q.Answer).ToList()));
            ReplaceText("[Q8Answer]", ConvertListToLines(currentIndividualReport.AllQuestions.Where(q => q.QuestionNumber == 8).Select(q => q.Answer).ToList()));

            for (int questionNumber = 9; questionNumber <= 12; questionNumber++)
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

            ReplaceText("[Q13Answer]", ConvertListToLines(currentIndividualReport.AllQuestions.Where(q => q.QuestionNumber == 13).Select(q => q.Answer).ToList()));

            currentIndividualReport.TotalPerformance = (double)currentIndividualReport.AllQuestions.Where(q => q.Answer == QuantitativeChoices.StronglyAgree || q.Answer == QuantitativeChoices.Agree).Count() / (currentIndividualReport.SurveyEntries.Count() * numberOfQuantitativeQuestions);
            ReplaceText("[Score]", GetPercentage(currentIndividualReport.TotalPerformance));
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

        private void WriteSummaryReport()
        {
            var stringBuilder = new StringBuilder();
            IList<string> header = new List<string>();
            if (surveyName == SurveyNames.StudentSurveyOnLecturer)
            {
                header.Add("Unit Code");
                header.Add("Unit Name");
                header.Add("Lecturer");
                header.Add("Overall Performance Indicator (%)");
                header.Add("Status");
                header.Add("Number of Flag");
                header.Add("Response Rate (%)");
                header.Add("Response");
                header.Add("Population");
                header.Add("Question 1");
                header.Add("Question 2");
                header.Add("Question 3");
                header.Add("Question 4");
                header.Add("Question 5");
                header.Add("Question 6");
                header.Add("Question 7");
                header.Add("Question 8");
                header.Add("Question 9");
                header.Add("Question 10");
            }
            else if (surveyName == SurveyNames.UnitAndLecturerSurvey)
            {
                header.Add("Unit Code");
                header.Add("Unit Name");
                header.Add("Lecturer");
                header.Add("Overall Performance Indicator (%)");
                header.Add("Status");
                header.Add("Number of Flag");
                header.Add("Response Rate (%)");
                header.Add("Response");
                header.Add("Population");
                header.Add("Question 1");
                header.Add("Question 2");
                header.Add("Question 3");
                header.Add("Question 4");
                header.Add("Question 5");
                header.Add("Question 6");
                header.Add("Question 9");
                header.Add("Question 10");
                header.Add("Question 11");
                header.Add("Question 12");
            }
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

                if (surveyName == SurveyNames.StudentSurveyOnLecturer)
                {
                    for (int questionNumber = 1; questionNumber <= 10; questionNumber++)
                    {
                        var performance = currentIndividualReport.AllQuestions.Where(q => q.QuestionNumber == questionNumber && (q.Answer == QuantitativeChoices.StronglyAgree || q.Answer == QuantitativeChoices.Agree)).Count();
                        body.Add(performance.ToString());
                    }
                }
                else
                {
                    for (int questionNumber = 1; questionNumber <= 6; questionNumber++)
                    {
                        var performance = currentIndividualReport.AllQuestions.Where(q => q.QuestionNumber == questionNumber && (q.Answer == QuantitativeChoices.StronglyAgree || q.Answer == QuantitativeChoices.Agree)).Count();
                        body.Add(performance.ToString());
                    }

                    for (int questionNumber = 9; questionNumber <= 12; questionNumber++)
                    {
                        var performance = currentIndividualReport.AllQuestions.Where(q => q.QuestionNumber == questionNumber && (q.Answer == QuantitativeChoices.StronglyAgree || q.Answer == QuantitativeChoices.Agree)).Count();
                        body.Add(performance.ToString());
                    }
                }
                stringBuilder.AppendLine(string.Join(",", body.ToArray()));
            }

            File.WriteAllText(Path.Combine(reportsDestination, "SummaryReport.csv"), stringBuilder.ToString());

            ProgressCompleted?.Invoke(this, "Summary report has been completed");
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
