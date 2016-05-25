using System;

namespace DocumentParser.common
{
    public class ResponseResult
    {
        string repCode;
        string repMessage;
        Object repData;

        public ResponseResult(string repCode, string repMessage, object repData)
        {
            this.repCode = repCode;
            this.repMessage = repMessage;
            this.repData = repData;
        }

        public ResponseResult()
        {
        }

        public string RepCode
        {
            get { return repCode; }
            set { repCode = value; }
        }

        public string RepMessage
        {
            get { return repMessage; }
            set { repMessage = value; }
        }

        public object RepData
        {
            get { return repData; }
            set { repData = value; }
        }
    }
}
