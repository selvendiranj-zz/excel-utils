using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excel_utils.Models
{
    public class MSGSetting
    {
        private string from = Environment.MachineName + "@company.com";
        private string to = "";
        private string cc = "";
        private string subject = "";
        private string body = "";
        private string attch = "";
        private string importance = "";

        public string From { get => from; set => from = value; }
        public string To { get => to; set => to = value; }
        public string Cc { get => cc; set => cc = value; }
        public string Subject { get => subject; set => subject = value; }
        public string Body { get => body; set => body = value; }
        public string Attch { get => attch; set => attch = value; }
        public string Importance { get => importance; set => importance = value; }
    }
}
