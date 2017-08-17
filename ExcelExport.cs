using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Common;
using System.Transactions;
using System.Data;
using Microsoft.VisualBasic;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Net;
using System.Net.Mail;
using System.Text.RegularExpressions;
using System.Net.Mime;
using System.Data.OleDb;

namespace SQLToXlsExport
{
    public class ExelExport
    {
        private Microsoft.Office.Interop.Excel.Application _application;
        private Microsoft.Office.Interop.Excel.Workbook _workbook;
        private Microsoft.Office.Interop.Excel.Worksheet _worksheet;

        //SQL Server Settings
        private string SQLServer = "";
        private string SQLDatabase = "";
        private string SQLConnString = "Data Source={0};Initial Catalog={1};Integrated Security=SSPI;Persist Security Info=False;Network Library=DBMSSOCN;Packet Size=8192;Max Pool Size=200;";
        private string SQLQuery = "";
        private string SQLQType = "Text";
        private string SQLErrFile = "";
        private string SQLParms = "";
        private object[] Parms = null;
        //Excel File Settings
        private string ExlConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties={2}Excel 8.0;HDR={1}{2}";
        private string ExlFileName = "";
        private string ExlHasHeader = "Yes";
        private string ExlSheets = "Sheet1";
        private string ExlFntName = "Arial";
        private bool ExlIsNew = true;
        private bool ExlDelRow = false;
        private int ExlFntSize = 10;
        private int ExlZoomPct = 100;
        private int ExlHdrPosn = 1;
        private bool ExlIsFormat = false;
        private string ExlFormat = "";
        //Mail settings
        private string MailFrom = System.Environment.MachineName + "@company.com";
        private string MailTo = "";
        private string MailCc = "";
        private string MailSubject = "";
        private string MailBody = "";
        private string MailAttch = "";
        private string MailImp = "";

        private string param = "";
        private DataAccess dataAccess = null;

        private ExelExport()
        {

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

        public ExelExport(string[] args)
        {
            try
            {
                for (int i = 0; i < args.Length; i++)
                {
                    param = args[i].Substring(1);

                    switch (param)
                    {
                        case "s": SQLServer = args[++i]; break;
                        case "d": SQLDatabase = args[++i]; break;
                        case "t": SQLQType = args[++i]; break;
                        case "q": SQLQuery = args[++i]; break;
                        case "p": SQLParms = args[++i]; break;
                        case "b": SQLErrFile = args[++i]; break;
                        case "o": ExlFileName = args[++i]; break;
                        case "e": ExlSheets = args[++i]; break;
                        case "h": ExlHdrPosn = Int32.Parse(args[++i]); break;
                        case "n": ExlIsNew = false; break;
                        case "m": ExlDelRow = true; break;
                        case "f": ExlFormat = args[++i]; ExlIsFormat = true; break;
                        case "mf": MailFrom = args[++i]; break;
                        case "mt": MailTo = args[++i]; break;
                        case "mc": MailCc = args[++i]; break;
                        case "ms": MailSubject = args[++i]; break;
                        case "mb": MailBody = args[++i]; break;
                        case "ma": MailAttch = args[++i]; break;
                        case "mp": MailImp = args[++i]; break;
                        case "?": PrintHelpDocument(); return;
                        default: new Exception("Can not parse input command line properly"); break;
                    }
                }
                if (ExlIsFormat)
                {
                    ExlFntName = ExlFormat.Split(',')[0];
                    ExlFntSize = Int32.Parse(ExlFormat.Split(',')[1]);
                    ExlZoomPct = Int32.Parse(ExlFormat.Split(',')[2]);
                }
                if (ExlHdrPosn < 0)
                {
                    ExlHasHeader = "No";
                    ExlHdrPosn = -ExlHdrPosn;
                }
                if (SQLServer != null && SQLDatabase != null && ExlFileName != null && !ExlFileName.Equals(string.Empty))
                {
                    if (ExlIsNew)
                    {
                        ExlHasHeader = "Yes";
                        ExlDelRow = false;
                        ExlHdrPosn = 1;
                    }
                    SQLConnString = string.Format(SQLConnString, SQLServer, SQLDatabase);
                    ExlConnString = string.Format(ExlConnString, ExlFileName, ExlHasHeader, ControlChars.Quote);

                    if (SQLQType.ToLower().Equals("procedure"))
                    {

                    }
                    else if (SQLQType.ToLower().Equals("text"))
                    {

                    }
                    else if (SQLQType.ToLower().Equals("file"))
                    {
                        StreamReader reader = new StreamReader(SQLQuery);
                        SQLQuery = reader.ReadToEnd();
                        SQLQType = "text";
                    }

                    else
                    {
                        new Exception("SQL Query Type is not specified correctly. valid values are -t Text/Procedure/File");
                    }

                    if (SQLParms != null && !SQLParms.Equals(string.Empty))
                    {
                        Parms = new object[SQLParms.Split(',').Count()];
                        int index = 0;
                        foreach (string sqlParm in SQLParms.Split(','))
                        {
                            if (sqlParm.ToLower().Equals("null"))
                            {
                                Parms[index] = null;
                            }
                            else
                            {
                                Parms[index] = sqlParm;
                            }
                            index++;
                        }
                    }

                    dataAccess = new DataAccess(SQLConnString, SQLQType);
                    DataSet ds = dataAccess.GetDataSet(SQLQuery, Parms, ExlSheets.Split(','));
                    TransferToExcel(ds);
                    FormatExcel(ExlFileName);
                }

                if (MailTo != null && !MailTo.Equals(string.Empty))
                {
                    if (MailFrom != null && MailFrom.ToLower().Equals(System.Environment.UserName.ToLower()))
                    {
                        MailFrom = MailFrom + "@company.com";
                    }
                    SendEMail(MailFrom, MailTo, MailCc, MailSubject, MailBody, MailAttch);
                }
            }
            catch (Exception ex)
            {
                if (SQLErrFile != null && !SQLErrFile.Equals(string.Empty))
                {
                    if (File.Exists(SQLErrFile))
                    {
                        File.Delete(SQLErrFile);
                    }
                    using (StreamWriter outfile = new StreamWriter(SQLErrFile))
                    {
                        outfile.Write(ex.InnerException);
                        foreach (string line in ex.Message.Split('\n'))
                        {
                            outfile.WriteLine(line);
                        }
                        outfile.Write(ex.StackTrace);
                        outfile.Close();
                    }
                }
                throw ex;
            }
            finally
            {
                //Console.WriteLine("Press any key to complete the process");
                //Console.ReadKey();
            }
        }

        private bool TransferToExcel(DataSet ds)
        {
            string strCreate = "";
            string strInsert = "";
            string strValues = "";

            foreach (System.Data.DataTable sheet in ds.Tables)
            {
                strCreate = "";
                strInsert = "";
                strValues = "";
                OleDbCommand cmd = new OleDbCommand();
                try
                {
                    cmd.Connection = new System.Data.OleDb.OleDbConnection(ExlConnString);
                    cmd.CommandTimeout = 300;
                    foreach (DataColumn column in sheet.Columns)
                    {
                        strCreate = strCreate + "[" + column.ColumnName + "] " + dataAccess.SQLType(column.DataType).ToString() + ",";
                        strInsert = strInsert + "[" + column.ColumnName + "],";
                        strValues = strValues + "?,";
                        cmd.Parameters.Add("@" + column.ColumnName, dataAccess.SQLType(column.DataType)).SourceColumn = column.ColumnName;
                    }
                    strCreate = strCreate.Remove(strCreate.Length - 1, 1);
                    strInsert = strInsert.Remove(strInsert.Length - 1, 1);
                    strValues = strValues.Remove(strValues.Length - 1, 1);
                    if (ExlIsNew)
                    {
                        cmd.Connection.Open();
                        cmd.CommandText = "CREATE TABLE [" + sheet.TableName + "] (" + strCreate + ")";
                        cmd.ExecuteNonQuery();
                        cmd.Connection.Close();
                    }

                    //cmd.CommandText = "INSERT INTO [" + sheet.TableName + ((ExlIsNew) ? "" : "$") + "](" + strInsert + ") VALUES (" + strValues + ")";
                    cmd.CommandText = "INSERT INTO [" + sheet.TableName + ((ExlIsNew) ? "" : "$") + "] VALUES (" + strValues + ")";
                    //Apply the dataset changes to the actual data source (the workbook).
                    //using (TransactionScope scope = new TransactionScope())
                    //{
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter())
                    {
                        adapter.InsertCommand = cmd;
                        int count = adapter.Update(sheet);
                    }
                    //scope.Complete();
                    cmd.Parameters.Clear();
                    //}
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    cmd.Dispose();
                }
            }
            ds.Dispose();
            return true;
        }

        private void ExportDatasetToExcel(DataSet ds, string strExcelFile)
        {
            System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection(ExlConnString);
            conn.Open();

            string[] strTableQ = new string[ds.Tables.Count + 1];

            int i = 0;

            //making table query

            for (i = 0; i <= ds.Tables.Count - 1; i++)
            {
                strTableQ[i] = "CREATE TABLE [" + ds.Tables[i].TableName + "](";

                int j = 0;
                for (j = 0; j <= ds.Tables[i].Columns.Count - 1; j++)
                {
                    DataColumn dCol = null;
                    dCol = ds.Tables[i].Columns[j];
                    strTableQ[i] += " [" + dCol.ColumnName + "] " + dataAccess.SQLType(dCol.DataType).ToString() + ",";
                }
                strTableQ[i] = strTableQ[i].Substring(0, strTableQ[i].Length - 2);
                strTableQ[i] += ")";

                OleDbCommand cmd = new OleDbCommand(strTableQ[i], conn);
                cmd.ExecuteNonQuery();
            }

            //making insert query
            string[] strInsertQ = new string[ds.Tables.Count];
            for (i = 0; i <= ds.Tables.Count - 1; i++)
            {
                strInsertQ[i] = "INSERT INTO " + ds.Tables[i].TableName + " Values (";
                for (int k = 0; k <= ds.Tables[i].Columns.Count - 1; k++)
                {
                    strInsertQ[i] += "@" + ds.Tables[i].Columns[k].ColumnName + " , ";
                }
                strInsertQ[i] = strInsertQ[i].Substring(0, strInsertQ[i].Length - 2);
                strInsertQ[i] += ")";
            }

            //Now inserting data
            for (i = 0; i <= ds.Tables.Count - 1; i++)
            {
                for (int j = 0; j <= ds.Tables[i].Rows.Count - 1; j++)
                {
                    OleDbCommand cmd = new OleDbCommand(strInsertQ[i], conn);
                    for (int k = 0; k <= ds.Tables[i].Columns.Count - 1; k++)
                    {
                        cmd.Parameters.AddWithValue("@" + ds.Tables[i].Columns[k].ColumnName.ToString(), ds.Tables[i].Rows[j][k].ToString());
                    }
                    cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();
                }
            }
        }

        private bool FormatExcel(string file)
        {
            try
            {
                double rowHeight = 16.0;
                _application = new Microsoft.Office.Interop.Excel.Application();
                _workbook = _application.Workbooks.Open(file, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                                Type.Missing, Type.Missing);
                string[] sheets = ExlSheets.Split(',');
                Array.Reverse(sheets);
                foreach (string sheet in sheets)
                {
                    _worksheet = (Microsoft.Office.Interop.Excel.Worksheet)_workbook.Sheets[((ExlIsNew) ? sheet.Replace(" ", "_") : sheet)];
                    _worksheet.Activate();

                    if (ExlHasHeader.Equals("Yes") && !ExlIsNew)
                    {
                        rowHeight = (double)_worksheet.Rows.get_Range(_worksheet.UsedRange.Cells[ExlHdrPosn, 1], _worksheet.UsedRange.Cells[ExlHdrPosn, _worksheet.UsedRange.Columns.Count]).Rows.EntireRow.RowHeight;
                    }
                    if (ExlIsFormat && ExlHasHeader.Equals("Yes") && _worksheet.UsedRange != null && _worksheet.UsedRange.Rows.Count > 0)
                    {
                        const String STYLE_NAME_HEADER = "HeaderFormat";
                        Microsoft.Office.Interop.Excel.Style sty1;
                        try
                        {
                            sty1 = _workbook.Styles[STYLE_NAME_HEADER];
                        }
                        catch
                        {
                            sty1 = _workbook.Styles.Add(STYLE_NAME_HEADER, Type.Missing);
                        }
                        sty1.Font.Name = ExlFntName;
                        sty1.Font.Size = ExlFntSize;
                        sty1.Font.Bold = true;
                        sty1.Font.Color = ColorTranslator.ToOle(Color.White);
                        sty1.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(51, 51, 51));
                        sty1.Borders.Color = ColorTranslator.ToOle(Color.White);
                        sty1.Borders.Weight = 1;
                        //Setting borders
                        sty1.Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
                        sty1.Borders[XlBordersIndex.xlDiagonalDown].LineStyle = XlLineStyle.xlLineStyleNone;
                        sty1.Borders[XlBordersIndex.xlDiagonalUp].LineStyle = XlLineStyle.xlLineStyleNone;
                        sty1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        sty1.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        sty1.WrapText = true;
                        //Apply header styles and freeze panes
                        _worksheet.Rows.get_Range(_worksheet.UsedRange.Cells[ExlHdrPosn, 1], _worksheet.UsedRange.Cells[ExlHdrPosn, _worksheet.UsedRange.Columns.Count]).Rows.Style = STYLE_NAME_HEADER;
                        _worksheet.get_Range(_worksheet.UsedRange.Cells[ExlHdrPosn + 1, 1], _worksheet.UsedRange.Cells[ExlHdrPosn + 1, 1]).Cells.Select();
                        _workbook.Windows[1].FreezePanes = true;
                    }
                    //FormatCondition dataFormat = _worksheet.Rows.get_Range(_worksheet.UsedRange.Cells[ExlHdrPosn + 1, 1], _worksheet.UsedRange.Cells[_worksheet.UsedRange.Rows.Count, _worksheet.UsedRange.Columns.Count]).FormatConditions.Add(XlFormatConditionType.xlExpression, XlFormatConditionOperator.xlEqual, Type.Missing, Type.Missing);
                    //dataFormat.Font.Name = ExlFntName;
                    if (ExlIsFormat)
                    {
                        _workbook.Windows[1].Zoom = ExlZoomPct;
                        _worksheet.Rows.get_Range(_worksheet.UsedRange.Cells[ExlHdrPosn + 1, 1], _worksheet.UsedRange.Cells[_worksheet.UsedRange.Rows.Count, _worksheet.UsedRange.Columns.Count]).EntireRow.Font.Name = ExlFntName;
                        _worksheet.Rows.get_Range(_worksheet.UsedRange.Cells[ExlHdrPosn + 1, 1], _worksheet.UsedRange.Cells[_worksheet.UsedRange.Rows.Count, _worksheet.UsedRange.Columns.Count]).EntireRow.Font.Size = ExlFntSize;
                    }

                    _worksheet.Rows.get_Range(_worksheet.UsedRange.Cells[ExlHdrPosn, 1], _worksheet.UsedRange.Cells[_worksheet.UsedRange.Rows.Count, _worksheet.UsedRange.Columns.Count]).EntireColumn.AutoFit();
                    _worksheet.Rows.get_Range(_worksheet.UsedRange.Cells[ExlHdrPosn, 1], _worksheet.UsedRange.Cells[_worksheet.UsedRange.Rows.Count, _worksheet.UsedRange.Columns.Count]).EntireRow.AutoFit();
                    _worksheet.Rows.get_Range(_worksheet.UsedRange.Cells[ExlHdrPosn, 1], _worksheet.UsedRange.Cells[ExlHdrPosn, _worksheet.UsedRange.Columns.Count]).Rows.EntireRow.RowHeight = rowHeight;
                    _worksheet.Rows.get_Range(_worksheet.UsedRange.Cells[ExlHdrPosn, 1], _worksheet.UsedRange.Cells[_worksheet.UsedRange.Rows.Count, _worksheet.UsedRange.Columns.Count]).EntireColumn.AutoFit();
                    //Alternate row formatting
                    if (ExlIsFormat)
                    {
                        FormatCondition format = _worksheet.Rows.get_Range(_worksheet.UsedRange.Cells[ExlHdrPosn + 1, 1], _worksheet.UsedRange.Cells[_worksheet.UsedRange.Rows.Count, _worksheet.UsedRange.Columns.Count]).FormatConditions.Add(XlFormatConditionType.xlExpression, XlFormatConditionOperator.xlEqual, "=MOD(ROW(),2) = 1", Type.Missing);
                        _workbook.set_Colors(28, ColorTranslator.ToOle(Color.AliceBlue));
                        format.Interior.Color = ColorTranslator.ToOle(Color.AliceBlue);
                        format.Interior.Pattern = XlPattern.xlPatternSolid;
                        format.Borders.Color = ColorTranslator.ToOle(Color.Silver);
                    }
                    if (ExlHasHeader.Equals("Yes") && ExlDelRow)
                    {
                        _worksheet.get_Range(_worksheet.UsedRange.Cells[ExlHdrPosn + 1, 1], _worksheet.UsedRange.Cells[ExlHdrPosn + 1, _worksheet.UsedRange.Columns.Count]).Delete(XlDeleteShiftDirection.xlShiftUp);
                    }
                    if (ExlHasHeader.Equals("No") && ExlDelRow)
                    {
                        _worksheet.get_Range(_worksheet.UsedRange.Cells[ExlHdrPosn, 1], _worksheet.UsedRange.Cells[ExlHdrPosn, _worksheet.UsedRange.Columns.Count]).Delete(XlDeleteShiftDirection.xlShiftUp);
                        _worksheet.get_Range(_worksheet.UsedRange.Cells[ExlHdrPosn, 1], _worksheet.UsedRange.Cells[ExlHdrPosn, _worksheet.UsedRange.Columns.Count]).Delete(XlDeleteShiftDirection.xlShiftUp);
                        _worksheet.get_Range(_worksheet.UsedRange.Cells[ExlHdrPosn, 1], _worksheet.UsedRange.Cells[ExlHdrPosn, 1]).Cells.Select();
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_worksheet);
                }
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

                if (_worksheet != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_worksheet);
                    _worksheet = null;
                }

                if (_workbook != null)
                {
                    _workbook.Close(true, Type.Missing, Type.Missing);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_workbook);
                    _workbook = null;
                }
                if (_application != null)
                {
                    _application.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_application);
                    _application = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private bool SendEMail(string from, string to, string cc, string subject, string body, string attachemnt)
        {
            try
            {
                var smtp = new SmtpClient
                {
                    Host = "10.10.242.16", //put your mail server address
                    Port = 25,
                    EnableSsl = false,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = true,
                    Credentials = new NetworkCredential()
                };
                using (var message = new MailMessage())
                {
                    System.Text.RegularExpressions.Match match;
                    string patternMailID = @"^[a-z][a-z|0-9|]*([_][a-z|0-9]+)*([.][a-z|"
                          + @"0-9]+([_][a-z|0-9]+)*)?@[a-z][a-z|0-9|]*\.([a-z]"
                          + @"[a-z|0-9]*(\.[a-z][a-z|0-9]*)?)$";
                    string patternFile = @"[a-zA-Z0-9]*.emsg";

                    match = Regex.Match(from, patternMailID, RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        message.From = new MailAddress(from);
                    }

                    foreach (string idTo in to.Split(';'))
                    {
                        match = Regex.Match(idTo, patternMailID, RegexOptions.IgnoreCase);
                        if (match.Success)
                        {
                            message.To.Add(idTo);
                        }
                        else
                        {
                            string tempMailid = GetMailId(idTo);
                            foreach (string idTof in tempMailid.Split(';'))
                            {
                                message.To.Add(idTof);
                            }
                        }
                    }
                    if (cc != null && !cc.Equals(string.Empty))
                    {
                        foreach (string idCc in cc.Split(';'))
                        {
                            match = Regex.Match(idCc, patternMailID, RegexOptions.IgnoreCase);
                            if (match.Success)
                            {
                                message.CC.Add(idCc);
                            }
                            else
                            {
                                string tempMailid = GetMailId(idCc);
                                foreach (string idCCf in tempMailid.Split(';'))
                                {
                                    message.CC.Add(idCCf);
                                }
                            }
                        }
                    }
                    message.Subject = subject;
                    if (Regex.Match(body, patternFile, RegexOptions.IgnoreCase).Success)
                    {
                        System.IO.StreamReader file = new System.IO.StreamReader(body);
                        body = file.ReadToEnd();
                    }

                    if (attachemnt != null && !attachemnt.Equals(string.Empty))
                    {
                        foreach (string file in attachemnt.Split(';'))
                        {
                            Attachment oAttch = new Attachment(file);
                            message.Attachments.Add(oAttch);
                        }
                    }
                    if (MailImp != null && !MailImp.Equals(string.Empty))
                    {
                        message.Priority = (MailPriority)Enum.Parse(typeof(MailPriority), MailImp);
                    }

                    //body = "<font face='Arial' size='10'>" + body + "</font>";

                    message.IsBodyHtml = true;
                    using (AlternateView altView = AlternateView.CreateAlternateViewFromString(body,
                    new ContentType(MediaTypeNames.Text.Html)))
                    {
                        message.AlternateViews.Add(altView);
                        smtp.Send(message);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.Message);
                throw ex;
            }
        }

        private string GetMailId(string fileName)
        {
            string mailId = "";

            string line;
            // Read the file and display it line by line.
            System.IO.StreamReader file = new System.IO.StreamReader(fileName);
            while ((line = file.ReadLine()) != null)
            {
                mailId = mailId + line + ";";
            }
            mailId = mailId.Remove(mailId.Length - 1, 1);
            file.Close();

            return mailId;
        }
    }
}
