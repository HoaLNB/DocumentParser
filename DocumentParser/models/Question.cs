using System.Collections.Generic;

namespace DocumentParser.models
{
    public class Question
    {
        private string qid;
        private Option questionContent;
        private List<Option> optionList;
        private string answer;
        private double mark;
        private string courseID;
        private string unit;
        private bool mixChoices;

        public Question(string qid, Option questionContent, List<Option> optionList, string answer, int mark,
            string courseId, string unit, bool mixChoices)
        {
            this.qid = qid;
            this.questionContent = questionContent;
            this.optionList = optionList;
            this.answer = answer;
            this.mark = mark;
            courseID = courseId;
            this.unit = unit;
            this.mixChoices = mixChoices;
        }

        /// <summary>
        /// Default constructor.
        /// </summary>
        public Question()
        {
            this.questionContent = new Option();
            this.optionList = new List<Option>();
//            for (int i = 0; i < 6; i++)
//            {
//                this.optionList.Add(new Option(string.Empty, string.Empty));
//            }
        }

        public void fillOptionList()
        {
            while (this.optionList.Count < 6)
            {
                this.optionList.Add(new Option(string.Empty, string.Empty));
            }
        }

        public string Qid
        {
            get { return qid; }
            set { qid = value; }
        }

        public Option QuestionContent
        {
            get { return questionContent; }
            set { questionContent = value; }
        }

        public List<Option> OptionList
        {
            get { return optionList; }
            set { optionList = value; }
        }

        public string Answer
        {
            get { return answer; }
            set { answer = value; }
        }

        public double Mark
        {
            get { return mark; }
            set { mark = value; }
        }

        public string CourseId
        {
            get { return courseID; }
            set { courseID = value; }
        }

        public string Unit
        {
            get { return unit; }
            set { unit = value; }
        }

        public bool MixChoices
        {
            get { return mixChoices; }
            set { mixChoices = value; }
        }
    }
}