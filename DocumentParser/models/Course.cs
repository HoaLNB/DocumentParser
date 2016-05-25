namespace DocumentParser.models
{
    public class Course
    {
        string courseID;
        string courseName;
        string description;
        int status;

        public Course(string courseId, string courseName, string description, int status)
        {
            courseID = courseId;
            this.courseName = courseName;
            this.description = description;
            this.status = status;
        }

        public Course(string courseId)
        {
            this.courseID = courseId;
            this.courseName = string.Empty;
            this.description = string.Empty;
            this.status = 1;
        }

        public string CourseId
        {
            get { return courseID; }
            set { courseID = value; }
        }

        public string CourseName
        {
            get { return courseName; }
            set { courseName = value; }
        }

        public string Description
        {
            get { return description; }
            set { description = value; }
        }

        public int Status
        {
            get { return status; }
            set { status = value; }
        }
    }
}