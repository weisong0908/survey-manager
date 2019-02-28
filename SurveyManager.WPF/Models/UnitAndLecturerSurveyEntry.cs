using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SurveyManager.WPF.Models
{
    public class UnitAndLecturerSurveyEntry: SurveyEntry
    {
        public UnitAndLecturerSurveyEntry(string[] columns) : base(columns)
        {
            for (int index = 1; index <= 6; index++)
                Questions.Add(new SurveyQuestion(index, ConvertQuantitativeAnswer(columns[index])));

            for (int index = 7; index <= 8; index++)
            {
                if (!string.IsNullOrEmpty(columns[index]) && !IsNoise(columns[index]))
                    Questions.Add(new SurveyQuestion(index, columns[index], isQuantitative: false));
            }

            for (int index = 9; index <= 12; index++)
                Questions.Add(new SurveyQuestion(index, ConvertQuantitativeAnswer(columns[index])));

            if (!string.IsNullOrEmpty(columns[13]) && !IsNoise(columns[13]))
                Questions.Add(new SurveyQuestion(13, columns[13], isQuantitative: false));
        }
    }
}
