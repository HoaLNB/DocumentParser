using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using DocumentParser.common;
using DocumentParser.models;

namespace DocumentParser.services
{
    /// <summary>
    /// This class includes methods to insert data into or get data from database.
    /// </summary>
    internal class DataAccessService
    {
        private static string connectionstring =
            @"Server=localhost;Database=QBank;Integrated security = true";
//        private static string connetionString = "Data Source=localhost,1433; Initial Catalog=QBank;User ID = sa; Password=sa123,Integrated security = false";

        /// <summary>
        ///     Get dataset from data from an entire table
        /// </summary>
        /// <returns></returns>
        public static DataSet selectDataSet(string tableName)
        {
            var connection = new SqlConnection(connectionstring);
            var dataAdapter = new SqlDataAdapter("select * from " + tableName, connection);
            var ds = new DataSet();
            dataAdapter.Fill(ds);
            return ds;
        }

        /// <summary>
        ///     Get dataset with custom query
        /// </summary>
        /// <param name="query"></param>
        /// <returns></returns>
        public static DataSet getCustomDataSet(string query)
        {
            var connection = new SqlConnection(connectionstring);
            connection.Open();
            var dataAdapter = new SqlDataAdapter(query, connection);
            var dataSet = new DataSet();
            dataAdapter.Fill(dataSet);
            connection.Close();
            return dataSet;
        }

        /// <summary>
        ///     Do command with custom query
        /// </summary>
        /// <param name="query"></param>
        public static void doCustomCommand(string query)
        {
            var connection = new SqlConnection(connectionstring);
            connection.Open();
            var command = new SqlCommand(query, connection);
            var reader = command.ExecuteReader();
            connection.Close();
        }

        /// <summary>
        ///     Insert a course object into Courses table
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public static bool insertCourse(Course i)
        {
            try
            {
                var conn = new SqlConnection(connectionstring);
                var cmd =
                    new SqlCommand(
                        "insert into "+Constants.DTB_COURSES + " values " + "(@" + Constants.PARAM_COURSE_COURSE_ID + ", @" +
                        Constants.PARAM_COURSE_COURSENAME + ", @" + Constants.PARAM_COURSE_DESCRIPTION + ", @" +
                        Constants.PARAM_COURSE_STATUS + ")",
                        conn);
                cmd.Parameters.AddWithValue("@" + Constants.PARAM_COURSE_COURSE_ID, i.CourseId);
                if (!i.CourseName.Equals(string.Empty)) cmd.Parameters.AddWithValue("@" + Constants.PARAM_COURSE_COURSENAME, i.CourseName);
                else cmd.Parameters.AddWithValue("@" + Constants.PARAM_COURSE_COURSENAME, DBNull.Value);

                if (!i.Description.Equals(string.Empty)) cmd.Parameters.AddWithValue("@" + Constants.PARAM_COURSE_DESCRIPTION, i.Description);
                else cmd.Parameters.AddWithValue("@" + Constants.PARAM_COURSE_DESCRIPTION, DBNull.Value);

                cmd.Parameters.AddWithValue("@" + Constants.PARAM_COURSE_STATUS, i.Status);
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Above: "+ex.Data);
                return false;
            }
        }

        /// <summary>
        ///     Insert a course object into Questions table. Spaghetti Code version.
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public static ResponseResult insertQuestion(Question q)
        {
            ResponseResult result = new ResponseResult();
            //Check to see if qid already exists.
            if (doesQidAlreadyExist(q.Qid))
            {
                result.RepCode = Constants.RC_SQL_INSERT_ONE_FAIL;
                return result;
            }
            try
            {
                var conn = new SqlConnection(connectionstring);
                var optionLetter = Constants.BEGIN_LETTER;
                var optionStringInQuery = string.Empty;
                while (optionLetter != Constants.AFTER_END_LETTER)
                {
                    optionStringInQuery += "@" + Constants.PARAM_QUESTION_OPT_FIRST + optionLetter + Constants.PARAM_QUESTION_OPT_TXT + ",@" + Constants.PARAM_QUESTION_OPT_FIRST + optionLetter + Constants.PARAM_QUESTION_OPT_IMG + ",";
                    optionLetter = StringUtils.alphabetIncrement(optionLetter);
                }
                var cmd = new SqlCommand("insert into " + Constants.DTB_QUESTIONS + " values " + "(@" + 
                                         Constants.PARAM_QUESTION_QID + ", @" + Constants.PARAM_QUESTION_QUESTION + ", @" +
                                         Constants.PARAM_QUESTION_IMAGE + ", " + optionStringInQuery + "@" +
                                         Constants.PARAM_QUESTION_ANSWER + ",@" + Constants.PARAM_QUESTION_MARK + ",@" +
                                         Constants.PARAM_COURSE_COURSE_ID + ",@" + Constants.PARAM_QUESTION_UNIT + ",@" +
                                         Constants.PARAM_QUESTION_MIXCHOICES + ")", conn);
                cmd.Parameters.AddWithValue("@" + Constants.PARAM_QUESTION_QID, q.Qid);
                cmd.Parameters.AddWithValue("@" + Constants.PARAM_QUESTION_QUESTION, q.QuestionContent.OptionText);
                cmd.Parameters.AddWithValue("@" + Constants.PARAM_QUESTION_IMAGE, q.QuestionContent.ImageLinkText);
                optionLetter = Constants.BEGIN_LETTER;
                foreach (var option in q.OptionList)
                {
                    cmd.Parameters.AddWithValue("@" + Constants.PARAM_QUESTION_OPT_FIRST + optionLetter + Constants.PARAM_QUESTION_OPT_TXT, option.OptionText);
                    cmd.Parameters.AddWithValue("@" + Constants.PARAM_QUESTION_OPT_FIRST + optionLetter + Constants.PARAM_QUESTION_OPT_IMG, option.ImageLinkText);
                    optionLetter = StringUtils.alphabetIncrement(optionLetter);
                }
                while (optionLetter != Constants.AFTER_END_LETTER)
                {
                    cmd.Parameters.AddWithValue("@" + Constants.PARAM_QUESTION_OPT_FIRST + optionLetter + Constants.PARAM_QUESTION_OPT_TXT, DBNull.Value);
                    cmd.Parameters.AddWithValue("@" + Constants.PARAM_QUESTION_OPT_FIRST + optionLetter + Constants.PARAM_QUESTION_OPT_IMG, DBNull.Value);
                    optionLetter = StringUtils.alphabetIncrement(optionLetter);
                }
                cmd.Parameters.AddWithValue("@" + Constants.PARAM_QUESTION_ANSWER, q.Answer);
                cmd.Parameters.AddWithValue("@" + Constants.PARAM_QUESTION_MARK, q.Mark);
                if (!isCourseIdAvailable(q.CourseId))
                {
                    insertCourse(new Course(q.CourseId));
                }
                cmd.Parameters.AddWithValue("@" + Constants.PARAM_COURSE_COURSE_ID, q.CourseId);
                cmd.Parameters.AddWithValue("@" + Constants.PARAM_QUESTION_UNIT, q.Unit);
                cmd.Parameters.AddWithValue("@" + Constants.PARAM_QUESTION_MIXCHOICES, q.MixChoices);
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();

                result.RepCode = Constants.RC_SQL_INSERT_ONE_SUCCESSFUL;
                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Data);
                result.RepCode = Constants.RC_SQL_INSERT_ONE_FAIL_UNKNOWN;
                return result;
            }
        }
        /// <summary>
        /// check if CourseID is available (to see if it's neccessary to add new Course to Courses table.
        /// </summary>
        /// <param name="courseID"></param>
        /// <returns></returns>
        public static bool isCourseIdAvailable(string courseID)
        {
            DataSet dataSet = getCustomDataSet(Constants.QUERY_CHECK_COURSEID + courseID +"'");
            return dataSet.Tables[0].Rows.Count != 0;
        }

        /// <summary>
        /// Check if qid already exist (to skip inserting).
        /// </summary>
        /// <param name="qid"></param>
        /// <returns></returns>
        public static bool doesQidAlreadyExist(string qid)
        {
            DataSet dataSet = getCustomDataSet(Constants.QUERY_CHECK_QID + qid + "'");
            return dataSet.Tables[0].Rows.Count != 0;
        }

        /// <summary>
        /// Read the database to get question list.
        /// </summary>
        /// <returns></returns>
        public static List<Question> getQuestionListFromDatabase()
        {
            List<Question> questionList = new List<Question>();
            DataSet questionDataSet =  selectDataSet(Constants.DTB_QUESTIONS);
            DataTable dataTable = questionDataSet.Tables[0];
            DataTableReader dataTableReader = dataTable.CreateDataReader();
            if (dataTableReader.HasRows)
            {
                while (dataTableReader.Read())
                {
                    Question question = new Question();
                    question.fillOptionList(); //to avoid null list
                    question.Qid = dataTableReader[Constants.PARAM_QUESTION_QID].ToString().Trim();
                    question.QuestionContent.OptionText = dataTableReader[Constants.PARAM_QUESTION_QUESTION].ToString().Trim();
                    question.QuestionContent.ImageLinkText = dataTableReader[Constants.PARAM_QUESTION_IMAGE].ToString().Trim();
                    question.OptionList[0].OptionText = dataTableReader[Constants.PARAM_QUESTION_OPTA_TXT].ToString().Trim();
                    question.OptionList[0].ImageLinkText = dataTableReader[Constants.PARAM_QUESTION_OPTA_IMG].ToString().Trim();
                    question.OptionList[1].OptionText = dataTableReader[Constants.PARAM_QUESTION_OPTB_TXT].ToString().Trim();
                    question.OptionList[1].ImageLinkText = dataTableReader[Constants.PARAM_QUESTION_OPTB_IMG].ToString().Trim();
                    question.OptionList[2].OptionText = dataTableReader[Constants.PARAM_QUESTION_OPTC_TXT].ToString().Trim();
                    question.OptionList[2].ImageLinkText = dataTableReader[Constants.PARAM_QUESTION_OPTC_IMG].ToString().Trim();
                    question.OptionList[3].OptionText = dataTableReader[Constants.PARAM_QUESTION_OPTD_TXT].ToString().Trim();
                    question.OptionList[3].ImageLinkText = dataTableReader[Constants.PARAM_QUESTION_OPTD_IMG].ToString().Trim();
                    question.OptionList[4].OptionText = dataTableReader[Constants.PARAM_QUESTION_OPTE_TXT].ToString().Trim();
                    question.OptionList[4].ImageLinkText = dataTableReader[Constants.PARAM_QUESTION_OPTE_IMG].ToString().Trim();
                    question.OptionList[5].OptionText = dataTableReader[Constants.PARAM_QUESTION_OPTF_TXT].ToString().Trim();
                    question.OptionList[5].ImageLinkText = dataTableReader[Constants.PARAM_QUESTION_OPTF_IMG].ToString().Trim();
                    question.Answer = dataTableReader[Constants.PARAM_QUESTION_ANSWER].ToString().Trim();
                    question.Mark = double.Parse(dataTableReader[Constants.PARAM_QUESTION_MARK].ToString().Trim());
                    question.CourseId = dataTableReader[Constants.PARAM_COURSE_COURSE_ID].ToString().Trim();
                    question.Unit = dataTableReader[Constants.PARAM_QUESTION_UNIT].ToString().Trim();
                    question.MixChoices = Constants.TRUE_TEXT.Equals(dataTableReader[Constants.PARAM_QUESTION_MIXCHOICES].ToString().Trim()) ? true:false;
                    questionList.Add(question);
                }
            }
            return questionList;
        }
    }
}