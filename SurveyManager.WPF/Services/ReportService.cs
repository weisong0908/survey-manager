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
        private readonly string surveyDataLocation;
        private readonly string reportDataLocation;
        private readonly string reportsDestination;
        private object reportTemplate;
        private IList<StudentSurveyOnLecturerSurveyEntry> surveyEntries;
        private IList<IndividualReport> individualReports;
        private Word.Application app;
        private Word.Document doc;

        public ReportService(string surveyDataLocation, string reportDataLocation, string reportsDestination, string reportTemplate)
        {
            this.surveyDataLocation = surveyDataLocation;
            this.reportDataLocation = reportDataLocation;
            this.reportsDestination = reportsDestination;
            this.reportTemplate = reportTemplate;
        }

        public void ReadSurveyData()
        {
            surveyEntries = new List<StudentSurveyOnLecturerSurveyEntry>();

            using (var parser = new TextFieldParser(surveyDataLocation))
            {
                parser.Delimiters = new string[] { "," };
                parser.HasFieldsEnclosedInQuotes = true;

                string[] headers = parser.ReadFields();

                while (!parser.EndOfData)
                {
                    string[] row = parser.ReadFields();

                    surveyEntries.Add(new StudentSurveyOnLecturerSurveyEntry(row));
                }
            }
        }

        public void ReadReportData()
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

        public void WriteIndividualReports()
        {
            app = new Word.Application();

            var numberOfQuantitativeQuestions = surveyEntries.First().Questions.Where(q => q.IsQuantitative == true).Count();

            foreach (var individualReport in individualReports)
            {
                doc = app.Documents.Add(ref reportTemplate);

                var entries = surveyEntries.Where(se => se.SurveyId == individualReport.SurveyId);
                var allQuestions = entries.SelectMany(r => r.Questions);

                ReplaceText("[StudyTerm]", individualReport.StudyTerm);
                ReplaceText("[LecturerName]", individualReport.Lecturer);
                ReplaceText("[Unit]", $"{individualReport.UnitCode.ToString()} - {individualReport.UnitName}");
                ReplaceText("[Population]", individualReport.ClassSize.ToString());
                ReplaceText("[Response]", $"{entries.Count().ToString()} ({GetPercent((double)entries.Count() / individualReport.ClassSize)})");

                for (int questionNumber = 1; questionNumber <= 10; questionNumber++)
                {
                    var numberOfStronglyAgree = allQuestions.Where(q => q.QuestionNumber == questionNumber && q.Answer == QuantitativeChoices.StronglyAgree).Count();
                    var numberOfAgree = allQuestions.Where(q => q.QuestionNumber == questionNumber && q.Answer == QuantitativeChoices.Agree).Count();
                    var numberOfNeutral = allQuestions.Where(q => q.QuestionNumber == questionNumber && q.Answer == QuantitativeChoices.Neutral).Count();
                    var numberOfDisagree = allQuestions.Where(q => q.QuestionNumber == questionNumber && q.Answer == QuantitativeChoices.Disagree).Count();
                    var numberOfStronglyDisagree = allQuestions.Where(q => q.QuestionNumber == questionNumber && q.Answer == QuantitativeChoices.StronglyDisagree).Count();
                    var numberOfSkipped = allQuestions.Where(q => q.QuestionNumber == questionNumber && q.Answer == QuantitativeChoices.Skipped).Count();

                    ReplaceText($"[Q{questionNumber}SA]", $"{numberOfStronglyAgree.ToString()} ({GetPercent((double)numberOfStronglyAgree / individualReport.ClassSize)})");
                    ReplaceText($"[Q{questionNumber}A]", $"{numberOfAgree.ToString()} ({GetPercent((double)numberOfAgree / individualReport.ClassSize)})");
                    ReplaceText($"[Q{questionNumber}N]", $"{numberOfNeutral.ToString()} ({GetPercent((double)numberOfNeutral / individualReport.ClassSize)})");
                    ReplaceText($"[Q{questionNumber}D]", $"{numberOfDisagree.ToString()} ({GetPercent((double)numberOfDisagree / individualReport.ClassSize)})");
                    ReplaceText($"[Q{questionNumber}SD]", $"{numberOfStronglyDisagree.ToString()} ({GetPercent((double)numberOfStronglyDisagree / individualReport.ClassSize)})");
                    ReplaceText($"[Q{questionNumber}S]", $"{numberOfSkipped.ToString()} ({GetPercent((double)numberOfSkipped / individualReport.ClassSize)})");
                }

                var strengths = allQuestions.Where(q => q.QuestionNumber == 11).Select(q => q.Answer).ToList();
                var suggestions = allQuestions.Where(q => q.QuestionNumber == 12).Select(q => q.Answer).ToList();

                ReplaceText("[Q11Answer]", ConvertListToLines(strengths));
                ReplaceText("[Q12Answer]", ConvertListToLines(suggestions));

                ReplaceText("[SA]", GetPercent((double)allQuestions.Where(q => q.Answer == QuantitativeChoices.StronglyAgree).Count() / (entries.Count() * numberOfQuantitativeQuestions)));
                ReplaceText("[A]", GetPercent((double)allQuestions.Where(q => q.Answer == QuantitativeChoices.Agree).Count() / (entries.Count() * numberOfQuantitativeQuestions)));
                ReplaceText("[N]", GetPercent((double)allQuestions.Where(q => q.Answer == QuantitativeChoices.Neutral).Count() / (entries.Count() * numberOfQuantitativeQuestions)));
                ReplaceText("[D]", GetPercent((double)allQuestions.Where(q => q.Answer == QuantitativeChoices.Disagree).Count() / (entries.Count() * numberOfQuantitativeQuestions)));
                ReplaceText("[SD]", GetPercent((double)allQuestions.Where(q => q.Answer == QuantitativeChoices.StronglyDisagree).Count() / (entries.Count() * numberOfQuantitativeQuestions)));
                ReplaceText("[S]", GetPercent((double)allQuestions.Where(q => q.Answer == QuantitativeChoices.Skipped).Count() / (entries.Count() * numberOfQuantitativeQuestions)));

                ReplaceText("[Score]", GetPercent((double)allQuestions.Where(q => q.Answer == QuantitativeChoices.StronglyAgree || q.Answer == QuantitativeChoices.Agree).Count() / (entries.Count() * numberOfQuantitativeQuestions)));

                SetFlags(individualReport, entries);

                if(individualReport.Flags.Count == 0)
                    ReplaceText("[Flag]", "No flag.");

                object filename = Path.Combine(reportsDestination, $"{individualReport.UnitCode} - {individualReport.Lecturer}.pdf");
                doc.SaveAs2(filename, Word.WdSaveFormat.wdFormatPDF);
                doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            }

            app.Quit();
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

        private string GetPercent(double value)
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

        private void SetFlags(IndividualReport individualReport, IEnumerable<StudentSurveyOnLecturerSurveyEntry> results)
        {
            if (results.Count() == 0)
                individualReport.Flags.Add("No response collected");

            if (individualReport.ClassSize < 10)
                individualReport.Flags.Add("The class size is less than 10 students.");

            if (((double)results.Count() / individualReport.ClassSize) < 0.2)
                individualReport.Flags.Add("The response rate is less than 20%.");
        }
    }
}
