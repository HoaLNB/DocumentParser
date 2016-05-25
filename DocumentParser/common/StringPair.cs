using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentParser.common
{
    public class StringPair
    {
        string firstString;
        string secondString;

        public StringPair(string firstString, string secondString)
        {
            this.firstString = firstString;
            this.secondString = secondString;
        }

        public string FirstString
        {
            get { return firstString; }
            set { firstString = value; }
        }

        public string SecondString
        {
            get { return secondString; }
            set { secondString = value; }
        }
    }
}
