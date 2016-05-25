using System.Collections.Generic;
using DocumentParser.common;
using DocumentParser.models;

namespace DocumentParser.services
{
    public static class QuestionService
    {
        public static List<Question> getQuestionListFromFile(string docxFilePath, string imgFolderPath)
        {
            DocumentFormatService wordUtilsOdk = new DocumentFormatService();
            List<Question> questionList = wordUtilsOdk.getQuestionListFromDOCX(docxFilePath, imgFolderPath);
            return questionList;
        }

        public static ResponseResult insertFromQuestionList(List<Question> questionList)
        {
            ResponseResult insertResult = new ResponseResult();
            List<ResponseResult> errorResponseList = new List<ResponseResult>();
            List<Question> errorQuestionList = new List<Question>();
            foreach (var question in questionList)
            {
                ResponseResult oneResponseResult = DataAccessService.insertQuestion(question);
                if (!Constants.RC_SQL_INSERT_ONE_SUCCESSFUL.Equals(oneResponseResult.RepCode))
                {
                    errorResponseList.Add(oneResponseResult);
                    errorQuestionList.Add(question);
                }
            }
            if (errorResponseList.Count > 0 && errorResponseList.Count < questionList.Count )
            {
                insertResult.RepCode = Constants.RC_SQL_INSERT_PARTIAL_PRIMARY_KEY;
                List<string> errorQidList = new List<string>();
                for (int i = 0; i < errorResponseList.Count; i++)
                {
                    errorQidList.Add(errorQuestionList[i].Qid);
                }
                insertResult.RepData = errorQidList;
            }else if (errorResponseList.Count > 0 && errorResponseList.Count == questionList.Count)
            {
                insertResult.RepCode = Constants.RC_SQL_INSERT_ALL_FAIL_PRIMARY_KEY;
            }else if (errorResponseList.Count == 0)
            {
                insertResult.RepCode = Constants.RC_SQL_INSERT_ALL_SUCCESSFUL;
            }
            return insertResult;
        }
    }
}
