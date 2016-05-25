using System.Collections.Generic;
using DocumentParser.common;

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
            questionContent = new Option();
            optionList = new List<Option>();
//            for (int i = 0; i < 6; i++)
//            {
//                this.optionList.Add(new Option(string.Empty, string.Empty));
//            }
        }

        public void fillOptionList()
        {
            while (optionList.Count < 6)
            {
                optionList.Add(new Option(string.Empty, string.Empty));
            }
        }

        public int countOptionExceptEmpty()
        {
            int numberOfOption = 0;
            foreach (var option in optionList)
            {
                if (!string.IsNullOrEmpty(option.OptionText))
                {
                    numberOfOption++;
                }
            }
            return numberOfOption;
        }

        /// <summary>
        /// Get list of elements to export.
        /// </summary>
        /// <returns></returns>
        public List<StringPair> getListOfElement()
        {
            List<StringPair> pairList = new List<StringPair>();
            if (!string.IsNullOrEmpty(QuestionContent.ImageLinkText))
            {
                pairList.Add(new StringPair(Constants.PATTERN_QID + qid,
                    QuestionContent.OptionText + "\r\n" + Constants.PATTERN_IMG + QuestionContent.ImageLinkText +
                    Constants.PATTERN_IMG_END_BRACKET));
            }
            else
            {
                pairList.Add(new StringPair(Constants.PATTERN_QID + qid, QuestionContent.OptionText));
            }
            char beginLetter = Constants.BEGIN_LETTER;
            foreach (var option in optionList)
            {
                if (!string.IsNullOrEmpty(option.OptionText))
                {
                    if (!string.IsNullOrEmpty(option.ImageLinkText))
                    {
                        pairList.Add(new StringPair(beginLetter + ".",
                            option.OptionText + "\r\n" + Constants.PATTERN_IMG + option.ImageLinkText +
                            Constants.PATTERN_IMG_END_BRACKET));
                    }
                    else
                    {
                        pairList.Add(new StringPair(beginLetter + ".", option.OptionText));
                    }
                    beginLetter = StringUtils.alphabetIncrement(beginLetter);
                }
            }
            pairList.Add(new StringPair(Constants.CELL_ANSWER, answer));
            pairList.Add(new StringPair(Constants.CELL_MARK, mark.ToString()));
            pairList.Add(new StringPair(Constants.CELL_COURSEID, courseID));
            pairList.Add(new StringPair(Constants.CELL_UNIT, unit));
            pairList.Add(new StringPair(Constants.CELL_MIXCHOICES, mixChoices? Constants.TRUE_TEXT: Constants.FALSE_TEXT));
            return pairList;
        }  

        public Question removeEmptyOptions()
        {
            Question returnQuestion = new Question();
            returnQuestion.qid = qid;
            returnQuestion.QuestionContent = QuestionContent;
            returnQuestion.Answer = Answer;
            returnQuestion.CourseId = courseID;
            returnQuestion.Mark = Mark;
            returnQuestion.Unit = Unit;
            returnQuestion.mixChoices = mixChoices;
            foreach (var option in optionList)
            {
                if (!string.IsNullOrEmpty(option.OptionText))
                {
                    returnQuestion.optionList.Add(new Option(option.OptionText,option.ImageLinkText));
                }
            }
            return returnQuestion;
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