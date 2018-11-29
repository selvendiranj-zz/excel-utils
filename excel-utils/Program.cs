using System;
using excel_utils.Models;

namespace excel_utils
{
    public class Program
    {
        private EMailClient eMailClient;
        private ExcelExport excelHelper;

        private static SQLSetting sql;
        private static XLSSetting xls;
        private static MSGSetting msg;

        private string param = string.Empty;

        public static void Main(string[] args)
        {
            sql = new SQLSetting();
            xls = new XLSSetting();
            msg = new MSGSetting();
        }

        /// <summary>
        /// Parse the parameters passed from console
        /// </summary>
        /// <param name="args"></param>
        private void ParseInput(string[] args)
        {
            try
            {
                for (int i = 0; i < args.Length; i++)
                {
                    param = args[i].Substring(1);

                    switch (param)
                    {
                        case "s": sql.Server = args[++i]; break;
                        case "d": sql.Database = args[++i]; break;
                        case "t": sql.QueryType = args[++i]; break;
                        case "q": sql.Query = args[++i]; break;
                        case "p": sql.Parms = args[++i]; break;
                        case "b": sql.ErrFile = args[++i]; break;
                        case "o": xls.FileName = args[++i]; break;
                        case "e": xls.Sheets = args[++i]; break;
                        case "h": xls.HdrPosn = Int32.Parse(args[++i]); break;
                        case "n": xls.IsNew = false; break;
                        case "m": xls.DelRow = true; break;
                        case "f": xls.Format = args[++i]; xls.IsFormat = true; break;
                        case "mf": msg.From = args[++i]; break;
                        case "mt": msg.To = args[++i]; break;
                        case "mc": msg.Cc = args[++i]; break;
                        case "ms": msg.Subject = args[++i]; break;
                        case "mb": msg.Body = args[++i]; break;
                        case "ma": msg.Attch = args[++i]; break;
                        case "mp": msg.Importance = args[++i]; break;
                        case "?": PrintHelpDocument(); return;
                        default: new Exception("Can not parse input command line properly"); break;
                    }
                }

                var xlClient = new ExcelExport(xls, sql);
                xlClient.ProcessExcel();

                var emailClient = new EMailClient(msg);
                emailClient.SendEMail();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
            finally
            {
                //Console.WriteLine("Press any key to complete the process");
                //Console.ReadKey();
            }
        }

        private void PrintHelpDocument()
        {
            string _help = "" + Environment.NewLine;
            _help += "[ -s: (Required) sql sever]" + Environment.NewLine;
            _help += "[ -d: (Required) sql database]" + Environment.NewLine;
            _help += "[ -t: (Optional) sql query type; default:Text; values:Text/Procedure/File]" + Environment.NewLine;
            _help += "[ -q: (Required) sql query text as SQLStatement/Procedure/FileName with path]" + Environment.NewLine;
            _help += "[ -p: (Optional) sql parameters if input sql is procedure]" + Environment.NewLine;
            _help += "[ -b: (Optional) sql error output fileName with path]" + Environment.NewLine;
            _help += "[ -o: (Required) excel output fileName with path]" + Environment.NewLine;
            _help += "[ -e: (Optional) excel sheet names; comma seperated values]" + Environment.NewLine;
            _help += "[ -h: (Optional) excel header position]" + Environment.NewLine;
            _help += "[ -n: (Optional) excel template use/new]" + Environment.NewLine;
            _help += "[ -m: (Optional) excel dummy first row deletion below header]" + Environment.NewLine;
            _help += "[ -f: (Optional) excel formatting values:font,size,zoom percent]" + Environment.NewLine;

            /*
            _help += "[-mf: (Optional) Mail from address; machineName if left as NULL]" + Environment.NewLine;
            _help += "[-mt: (Required) Mail to address; mailId/prm fileName with path]" + Environment.NewLine;
            _help += "[-mc: (Optional) Mail cc address; mailid/prm fileName with path]" + Environment.NewLine;
            _help += "[-ms: (Required) Mail Subject]" + Environment.NewLine;
            _help += "[-mb: (Required) Mail Body values:Text/emsg fileName with path]" + Environment.NewLine;
            _help += "[-ma: (Optional) Mail Attachments; fileName with path]" + Environment.NewLine;
            _help += "[-mp: (Optional) Mail Priority values: Low, Normal, High]" + Environment.NewLine;
            */

            _help += "[ -?:  Help" + Environment.NewLine;

            Console.WriteLine(_help);
        }
    }
}
