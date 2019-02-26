using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SurveyManager.WPF.Models
{
    public class StudentSurveyOnLecturerSurveyEntry : SurveyEntry
    {
        public StudentSurveyOnLecturerSurveyEntry(string[] columns) : base(columns)
        {
            for (int index = 1; index <= 10; index++)
            {
                Questions.Add(new SurveyQuestion(index, ConvertQuantitativeAnswer(columns[index])));
            }

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
    }
}
