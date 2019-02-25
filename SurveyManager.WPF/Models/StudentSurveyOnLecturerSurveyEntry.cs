using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SurveyManager.WPF.Models
{
    public class StudentSurveyOnLecturerSurveyEntry
    {
        public int SurveyId { get; set; }
        public IList<SurveyQuestion> Questions { get; set; }

        private IList<string> temporaryAnswers;

        public StudentSurveyOnLecturerSurveyEntry(string[] columns)
        {
            Questions = new List<SurveyQuestion>();
            temporaryAnswers = new List<string>();

            SurveyId = int.TryParse(columns[0], out int id) ? id : 0;

            for (int index = 1; index <= 10; index++)
            {
                Questions.Add(new SurveyQuestion(index, ConvertQuantitativeAnswer(columns[index])));
            }

            temporaryAnswers.Clear();
            for (int index = 11; index <= 13; index++)
            {
                if (!string.IsNullOrEmpty(columns[index]) && !IsNoise(columns[index]))
                    temporaryAnswers.Add(columns[index]);
            }
            Questions.Add(new SurveyQuestion(11, string.Join(".", temporaryAnswers.ToList()), isQuantitative: false));

            temporaryAnswers.Clear();
            for (int index = 14; index <= 16; index++)
            {
                if (!string.IsNullOrEmpty(columns[index]) && !IsNoise(columns[index]))
                    temporaryAnswers.Add(columns[index]);
            }
            Questions.Add(new SurveyQuestion(12, string.Join(".", temporaryAnswers.ToList()), isQuantitative: false));
        }

        private string ConvertQuantitativeAnswer(string value)
        {
            switch (value.ToLower())
            {
                case "strongly agree":
                case "stronglyagree":
                    return QuantitativeChoices.StronglyAgree;
                case "agree":
                    return QuantitativeChoices.Agree;
                case "not sure":
                case "neutral":
                    return QuantitativeChoices.Neutral;
                case "disagree":
                    return QuantitativeChoices.Disagree;
                case "strongly disagree":
                case "stronglydisagree":
                    return QuantitativeChoices.StronglyDisagree;
                default:
                    return QuantitativeChoices.Skipped;
            }
        }

        private bool IsNoise(string value)
        {
            switch (value.ToLower())
            {
                case "nil":
                case "na":
                case "no":
                case "-":
                    return true;
                default:
                    return false;
            }
        }
    }
}
