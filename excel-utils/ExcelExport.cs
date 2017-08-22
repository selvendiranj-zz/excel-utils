using System;
using System.Linq;
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
using excel_utils;
using excel_utils.Models;

namespace excel_utils
{
    public class ExelExport
    {
        private Application _application;
        private Workbook _workbook;
        private Worksheet _worksheet;
        private DataAccess dataAccess = null;

        private XLSSetting xls;
        private SQLSetting sql;

        public ExelExport(XLSSetting xls, SQLSetting sql)
        {
            this.xls = xls;
            this.sql = sql;
        }

        public void ProcessExcel()
        {
            try
            {
                if (xls.IsFormat)
                {
                    xls.FntName = xls.Format.Split(',')[0];
                    xls.FntSize = Int32.Parse(xls.Format.Split(',')[1]);
                    xls.ZoomPct = Int32.Parse(xls.Format.Split(',')[2]);
                }
                if (xls.HdrPosn < 0)
                {
                    xls.HasHeader = "No";
                    xls.HdrPosn = -xls.HdrPosn;
                }
                if (sql.Server != null && sql.Database != null &&
                        !string.IsNullOrEmpty(xls.FileName))
                {

                }

                if (xls.IsNew)
                {
                    xls.HasHeader = "Yes";
                    xls.DelRow = false;
                    xls.HdrPosn = 1;
                }
                sql.ConnString = string.Format(sql.ConnString, sql.Server, sql.Database);
                xls.ConnString = string.Format(xls.ConnString, xls.FileName, xls.HasHeader, ControlChars.Quote);

                if (sql.QueryType.ToLower().Equals("procedure"))
                {

                }
                else if (sql.QueryType.ToLower().Equals("text"))
                {

                }
                else if (sql.QueryType.ToLower().Equals("file"))
                {
                    StreamReader reader = new StreamReader(sql.Query);
                    sql.Query = reader.ReadToEnd();
                    sql.QueryType = "Text";
                }

                else
                {
                    new Exception("sql. Query Type is not specified correctly." +
                            " valid values are -t Text/Procedure/File");
                }

                if (sql.Parms != null && !sql.Parms.Equals(string.Empty))
                {
                    sql.ParmCollection = new object[sql.Parms.Split(',').Count()];
                    int index = 0;
                    foreach (string sqlParm in sql.Parms.Split(','))
                    {
                        if (sqlParm.ToLower().Equals("null"))
                        {
                            sql.ParmCollection[index] = null;
                        }
                        else
                        {
                            sql.ParmCollection[index] = sqlParm;
                        }
                        index++;
                    }
                }

                dataAccess = new DataAccess(sql.ConnString, sql.QueryType);
                DataSet ds = dataAccess.GetDataSet(sql.Query, sql.ParmCollection, xls.Sheets.Split(','));
                TransferToExcel(ds);
                FormatExcel(xls.FileName);
            }
            catch (Exception ex)
            {
                if (sql.ErrFile != null && !sql.ErrFile.Equals(string.Empty))
                {
                    if (File.Exists(sql.ErrFile))
                    {
                        File.Delete(sql.ErrFile);
                    }
                    using (StreamWriter outfile = new StreamWriter(sql.ErrFile))
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
                throw;
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
                    cmd.Connection = new OleDbConnection(xls.ConnString);
                    cmd.CommandTimeout = 300;

                    foreach (DataColumn column in sheet.Columns)
                    {
                        strCreate = strCreate + "[" + column.ColumnName + "] " +
                                    dataAccess.SQLType(column.DataType).ToString() + ",";
                        strInsert = strInsert + "[" + column.ColumnName + "],";
                        strValues = strValues + "?,";

                        cmd.Parameters
                           .Add("@" + column.ColumnName, dataAccess.SQLType(column.DataType))
                           .SourceColumn = column.ColumnName;
                    }

                    strCreate = strCreate.Remove(strCreate.Length - 1, 1);
                    strInsert = strInsert.Remove(strInsert.Length - 1, 1);
                    strValues = strValues.Remove(strValues.Length - 1, 1);

                    if (xls.IsNew)
                    {
                        cmd.Connection.Open();
                        cmd.CommandText = string.Format("CREATE TABLE [{0}] ({1})", sheet.TableName, strCreate);
                        cmd.ExecuteNonQuery();
                        cmd.Connection.Close();
                    }

                    //cmd.CommandText = "INSERT INTO [" + 
                    //  sheet.TableName + ((xls.IsNew) ? "" : "$") + "](" + strInsert + ") VALUES (" + strValues + ")";
                    string tableName = sheet.TableName + ((xls.IsNew) ? "" : "$");
                    cmd.CommandText = string.Format("INSERT INTO [{0}] VALUES ({1})", tableName, strValues);

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
            OleDbConnection conn = new OleDbConnection(xls.ConnString);
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
                    strTableQ[i] += " [" + dCol.ColumnName + "] " +
                        dataAccess.SQLType(dCol.DataType).ToString() + ",";
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
                        cmd.Parameters.AddWithValue(
                            "@" + ds.Tables[i].Columns[k].ColumnName.ToString(),
                            ds.Tables[i].Rows[j][k].ToString());
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
                _application = new Application();
                _workbook = _application.Workbooks.Open(file, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing);
                string[] sheets = xls.Sheets.Split(',');
                Array.Reverse(sheets);
                foreach (string sheet in sheets)
                {
                    _worksheet = (Worksheet)_workbook.Sheets[((xls.IsNew) ? sheet.Replace(" ", "_") : sheet)];
                    _worksheet.Activate();

                    if (xls.HasHeader.Equals("Yes") && !xls.IsNew)
                    {
                        rowHeight = (double)_worksheet.Rows.get_Range(
                            _worksheet.UsedRange.Cells[xls.HdrPosn, 1],
                            _worksheet.UsedRange.Cells[xls.HdrPosn,
                            _worksheet.UsedRange.Columns.Count]).Rows.EntireRow.RowHeight;
                    }
                    if (xls.IsFormat && xls.HasHeader.Equals("Yes") && _worksheet.UsedRange != null &&
                            _worksheet.UsedRange.Rows.Count > 0)
                    {
                        const String STYLE_NAME_HEADER = "HeaderFormat";
                        Style sty1;
                        try
                        {
                            sty1 = _workbook.Styles[STYLE_NAME_HEADER];
                        }
                        catch
                        {
                            sty1 = _workbook.Styles.Add(STYLE_NAME_HEADER, Type.Missing);
                        }
                        sty1.Font.Name = xls.FntName;
                        sty1.Font.Size = xls.FntSize;
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
                        _worksheet.Rows.get_Range(_worksheet.UsedRange.Cells[xls.HdrPosn, 1],
                            _worksheet.UsedRange.Cells[xls.HdrPosn,
                            _worksheet.UsedRange.Columns.Count]).Rows.Style = STYLE_NAME_HEADER;
                        _worksheet.get_Range(_worksheet.UsedRange.Cells[xls.HdrPosn + 1, 1],
                            _worksheet.UsedRange.Cells[xls.HdrPosn + 1, 1]).Cells.Select();
                        _workbook.Windows[1].FreezePanes = true;
                    }
                    //FormatCondition dataFormat = _worksheet.Rows.get_Range(_worksheet.UsedRange.Cells[xls.HdrPosn + 1, 1], _worksheet.UsedRange.Cells[_worksheet.UsedRange.Rows.Count, _worksheet.UsedRange.Columns.Count]).FormatConditions.Add(XlFormatConditionType.xlExpression, XlFormatConditionOperator.xlEqual, Type.Missing, Type.Missing);
                    //dataFormat.Font.Name = xls.FntName;
                    if (xls.IsFormat)
                    {
                        _workbook.Windows[1].Zoom = xls.ZoomPct;
                        _worksheet.Rows.get_Range(_worksheet.UsedRange.Cells[xls.HdrPosn + 1, 1],
                            _worksheet.UsedRange.Cells[_worksheet.UsedRange.Rows.Count,
                            _worksheet.UsedRange.Columns.Count]).EntireRow.Font.Name = xls.FntName;
                        _worksheet.Rows.get_Range(_worksheet.UsedRange.Cells[xls.HdrPosn + 1, 1],
                            _worksheet.UsedRange.Cells[_worksheet.UsedRange.Rows.Count,
                            _worksheet.UsedRange.Columns.Count]).EntireRow.Font.Size = xls.FntSize;
                    }

                    _worksheet.Rows.get_Range(_worksheet.UsedRange.Cells[xls.HdrPosn, 1],
                        _worksheet.UsedRange.Cells[_worksheet.UsedRange.Rows.Count,
                        _worksheet.UsedRange.Columns.Count]).EntireColumn.AutoFit();
                    _worksheet.Rows.get_Range(_worksheet.UsedRange.Cells[xls.HdrPosn, 1],
                        _worksheet.UsedRange.Cells[_worksheet.UsedRange.Rows.Count,
                        _worksheet.UsedRange.Columns.Count]).EntireRow.AutoFit();
                    _worksheet.Rows.get_Range(_worksheet.UsedRange.Cells[xls.HdrPosn, 1],
                        _worksheet.UsedRange.Cells[xls.HdrPosn,
                        _worksheet.UsedRange.Columns.Count]).Rows.EntireRow.RowHeight = rowHeight;
                    _worksheet.Rows.get_Range(_worksheet.UsedRange.Cells[xls.HdrPosn, 1],
                        _worksheet.UsedRange.Cells[_worksheet.UsedRange.Rows.Count,
                        _worksheet.UsedRange.Columns.Count]).EntireColumn.AutoFit();

                    //Alternate row formatting
                    if (xls.IsFormat)
                    {
                        FormatCondition format = _worksheet.Rows.get_Range(
                            _worksheet.UsedRange.Cells[xls.HdrPosn + 1, 1],
                            _worksheet.UsedRange.Cells[_worksheet.UsedRange.Rows.Count,
                            _worksheet.UsedRange.Columns.Count])
                            .FormatConditions.Add(XlFormatConditionType.xlExpression,
                                XlFormatConditionOperator.xlEqual, "=MOD(ROW(),2) = 1", Type.Missing);
                        _workbook.set_Colors(28, ColorTranslator.ToOle(Color.AliceBlue));
                        format.Interior.Color = ColorTranslator.ToOle(Color.AliceBlue);
                        format.Interior.Pattern = XlPattern.xlPatternSolid;
                        format.Borders.Color = ColorTranslator.ToOle(Color.Silver);
                    }
                    if (xls.HasHeader.Equals("Yes") && xls.DelRow)
                    {
                        _worksheet.get_Range(_worksheet.UsedRange.Cells[xls.HdrPosn + 1, 1],
                            _worksheet.UsedRange.Cells[xls.HdrPosn + 1,
                            _worksheet.UsedRange.Columns.Count]).Delete(XlDeleteShiftDirection.xlShiftUp);
                    }
                    if (xls.HasHeader.Equals("No") && xls.DelRow)
                    {
                        _worksheet.get_Range(_worksheet.UsedRange.Cells[xls.HdrPosn, 1],
                            _worksheet.UsedRange.Cells[xls.HdrPosn,
                            _worksheet.UsedRange.Columns.Count]).Delete(XlDeleteShiftDirection.xlShiftUp);
                        _worksheet.get_Range(_worksheet.UsedRange.Cells[xls.HdrPosn, 1],
                            _worksheet.UsedRange.Cells[xls.HdrPosn,
                            _worksheet.UsedRange.Columns.Count]).Delete(XlDeleteShiftDirection.xlShiftUp);
                        _worksheet.get_Range(_worksheet.UsedRange.Cells[xls.HdrPosn, 1],
                            _worksheet.UsedRange.Cells[xls.HdrPosn, 1]).Cells.Select();
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
    }
}
