using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SurveyManager.WPF.Models
{
    public class SurveyQuestion
    {
        public int QuestionNumber { get; set; }
        public string Answer { get; set; }
        public bool IsQuantitative { get; set; }

        public SurveyQuestion(int questionNumber, string answer, bool isQuantitative = true)
        {
            QuestionNumber = questionNumber;
            Answer = answer;
            IsQuantitative = isQuantitative;
        }
    }
}
