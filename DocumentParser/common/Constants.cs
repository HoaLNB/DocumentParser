namespace DocumentParser.common
{
    public class Constants
    {
        public static char BEGIN_LETTER = 'A';
        public static char AFTER_END_LETTER = 'G';
        public static int QUESTION_TABLE_ROW_EXCEPT_OPTIONS = 6;
        public static int QUESTION_TABLE_COLUMN = 2;
        public static string LINE_BREAK_CHARACTER = "\u2028\n";

        //Export table properties
        public static int QUESTION_TABLE_BORDER = 1;

        //Database name
        public static string DTB_COURSES = "Courses";
        public static string DTB_QUESTIONS = "Questions";

        //SQL Parameter
        public static string PARAM_COURSE_COURSE_ID = "CourseID";
        public static string PARAM_COURSE_COURSENAME = "CourseName";
        public static string PARAM_COURSE_DESCRIPTION = "Description";
        public static string PARAM_COURSE_STATUS = "Status";

        public static string PARAM_QUESTION_QID = "QID";
        public static string PARAM_QUESTION_QUESTION = "Question";
        public static string PARAM_QUESTION_IMAGE = "Image";
        public static string PARAM_QUESTION_ANSWER = "Answers";
        public static string PARAM_QUESTION_MARK = "Mark";
        public static string PARAM_QUESTION_UNIT = "Unit";
        public static string PARAM_QUESTION_MIXCHOICES = "MixChoices";
        public static string PARAM_QUESTION_OPT_FIRST = "Opt";
        public static string PARAM_QUESTION_OPT_TXT = "_Txt";
        public static string PARAM_QUESTION_OPT_IMG = "_Img";

        public static string PARAM_QUESTION_OPTA_TXT = "OptA_Txt";
        public static string PARAM_QUESTION_OPTA_IMG = "OptA_Img";
        public static string PARAM_QUESTION_OPTB_TXT = "OptB_Txt";
        public static string PARAM_QUESTION_OPTB_IMG = "OptB_Img";
        public static string PARAM_QUESTION_OPTC_TXT = "OptC_Txt";
        public static string PARAM_QUESTION_OPTC_IMG = "OptC_Img";
        public static string PARAM_QUESTION_OPTD_TXT = "OptD_Txt";
        public static string PARAM_QUESTION_OPTD_IMG = "OptD_Img";
        public static string PARAM_QUESTION_OPTE_TXT = "OptE_Txt";
        public static string PARAM_QUESTION_OPTE_IMG = "OptE_Img";
        public static string PARAM_QUESTION_OPTF_TXT = "OptF_Txt";
        public static string PARAM_QUESTION_OPTF_IMG = "OptF_Img";

        public static string QUERY_CHECK_COURSEID = "select 1 from Courses where CourseID = '";
        public static string QUERY_CHECK_QID = "select 1 from Questions where Qid = '";
        public static string PATTERN_IMG = @"[file:"; 
        public static char PATTERN_IMG_END_BRACKET = ']';
        public static string PATTERN_QID = "QN=";

        //Cell keyword
        public static string CELL_ANSWER = "ANSWER:";
        public static string CELL_MARK = "MARK:";
        public static string CELL_UNIT = "UNIT:";
        public static string CELL_COURSEID = "COURSE:";
        public static string CELL_MIXCHOICES = "MIX CHOICES:";

        public static string YES_CHOICE = "Yes";
        public static string NO_CHOICE = "No";
        public static string TRUE_TEXT = "True";
        public static string FALSE_TEXT = "False";

        public static string WORD_FILE_DIALOG_FILTER = @"Word 2007 or later document|*.docx";

        //Response code
        public static string RC_SQL_INSERT_ONE_SUCCESSFUL = "RC0000";
        public static string RC_SQL_INSERT_ONE_FAIL = "RC0001";
        public static string RC_SQL_INSERT_ONE_FAIL_UNKNOWN = "RC0002";
        public static string RC_SQL_INSERT_ALL_SUCCESSFUL = "RC0011";
        public static string RC_SQL_INSERT_PARTIAL_PRIMARY_KEY = "RC0012";
        public static string RC_SQL_INSERT_ALL_FAIL_PRIMARY_KEY = "RC0013";
        public static string RC_SQL_INSERT_ALL_FAIL_UNKNOWN = "RC0014";
        public static string RC_QUESTION_NOT_FOUND = "RC0015";

        //Response message


        //Message box
        public static string MSG_MB_INSERT_ALL_SUCCESSFUL = "All questions were successfully inserted!";
        public static string MSG_MB_INSERT_PARTIAL_FAIL = "The following questions were not inserted:\n";
        public static string MSG_MB_INSERT_PARTIAL_FAIL_AFTER = "\nPlease check the docx file again.";
        public static string MSG_MB_INSERT_ALL_FAIL = "Insert failed! No question was inserted!";
        public static string MSG_MB_NO_QUESTION = "No question was found in this file.";
        

    }
}
