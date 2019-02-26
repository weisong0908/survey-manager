using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SurveyManager.WPF.Models
{
    public class SurveyEntry
    {
        public int SurveyId { get; set; }
        public IList<SurveyQuestion> Questions { get; set; }
        protected IList<string> temporaryAnswers;

        public SurveyEntry(string[] columns)
        {
            Questions = new List<SurveyQuestion>();
            SurveyId = int.TryParse(columns[0], out int id) ? id : 0;
            temporaryAnswers = new List<string>();
        }

        protected string ConvertQuantitativeAnswer(string value)
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

        protected bool IsNoise(string value)
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
