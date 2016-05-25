using System.Collections.Generic;
using System.Windows.Forms;
using DocumentParser.models;
using DocumentParser.services;

namespace DocumentParser
{
    public static class Test
    {
        public static void testInsert()
        {
            //var course = new Course("ISC303", "E-Commerce", "Electronic Commercial", true);
//            var course = new Course("ISC303", "E-Commerce", "Electronic Commercial", true);
            var course = new Course("ISC305");
            DataAccessService.insertCourse(course);
            int i = 80;
            while (i < 85)
            {
                List<Option> optionList = new List<Option>();
                optionList.Add(new Option("Forever alone", "c0003"));
                optionList.Add(new Option("Forever in love", "c0005"));
                optionList.Add(new Option("Forever in solitude", "c0006"));
                Question question = new Question("00001" + i, new Option("How are you?", "00001.jpg"), optionList, "A", 2, "ISC301", "12", true);
                DataAccessService.insertQuestion(question);
                i++;
            }
            MessageBox.Show("Testing done!");
        }
    }
}