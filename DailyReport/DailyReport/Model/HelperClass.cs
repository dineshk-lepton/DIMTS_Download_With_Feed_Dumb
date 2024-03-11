using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DailyReport.Model
{
    public static class HelperClass
    {
        public static bool IsColumnExist(DataTable dt, string columnName, bool create = false)
        {
            bool isExist = false;
            try
            {
                string dtColumnName = ",";
                foreach (DataColumn column in dt.Columns)
                {
                    dtColumnName = dtColumnName + column.Caption.ToLower().Trim() + ",";
                }
                string[] columnList = columnName.Split(',');
                foreach (string column in columnList)
                {
                    if (!dtColumnName.Contains("," + column.ToLower().Trim() + ","))
                    {
                        if (create.Equals(true))
                        {
                            dt.Columns.Add(column.Trim(), typeof(String));
                        }
                        else
                        {
                            return false;
                        }
                    }
                }
                isExist = true;
            }
            catch (Exception ex)
            {

                isExist = false;
            }
            return isExist;
        }

        public static string ExportOutput(DataTable dtSource, string outputFilePath, string outputFileName, string extension, int exportSheetType = 0, bool exportToCsv = false, bool isErrorOutput = false, bool allLeftAlign = false)
        {

            exportToCsv = extension.Contains("csv") ? true : false;
            bool status = false;
            string outputPath = outputFilePath;

            if (exportToCsv)
            {
                outputPath = outputPath == string.Empty ? Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads\\" + outputFileName + "_" + DateTime.Now.ToString("HHmmss") + ".csv" : outputPath + "\\" + outputFileName + "_" + DateTime.Now.ToString("HHmmss") + ".csv";
                status = WriteCsv(dtSource, outputPath);
            }
            else
            {
                string[] files = Directory.GetFiles(outputFilePath);
                foreach (string file in files)
                { File.Delete(file); }

                string renderedContent = DatatableToExcel(dtSource, exportSheetType, allLeftAlign, isErrorOutput);
                if (!renderedContent.Equals(string.Empty))
                {
                    outputPath = outputPath == string.Empty ? Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads\\" + outputFileName + "_" + DateTime.Now.ToString("HHmmss") + ".xls" : outputPath + "\\" + outputFileName + "_" + DateTime.Now.ToString("HHmmss") + ".xls";
                    //outputPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads\\" + outputFileName + "_" + DateTime.Now.ToString("HHmmss") + ".xls";
                    File.WriteAllText(outputPath, renderedContent);
                    status = true;
                }
            }
            if (status)
            {
                //Process.Start(outputPath);
            }
            else
            {

            }
            return outputPath;
        }
        public static bool WriteCsv(DataTable dt, string destination)
        {
            bool status = false;
            try
            {
                using (var writer = new StreamWriter(destination))
                {
                    writer.WriteLine(string.Join(",", dt.Columns.Cast<DataColumn>().Select(dc => dc.ColumnName)));
                    foreach (DataRow row in dt.Rows)
                    {
                        writer.WriteLine(string.Join(",", row.ItemArray));
                    }
                }
                dt.Dispose();
                status = true;
            }
            catch (Exception ex)
            {

            }

            return status;
        }
        public static bool WriteExcel(DataSet dstable, string destination)
        {
            bool status = false;
            try
            {

                using (var workbook = SpreadsheetDocument.Create(destination, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
                {
                    var workbookPart = workbook.AddWorkbookPart();
                    workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
                    workbook.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();

                    uint sheetId = 1;
                    //DataTable table = dstable;
                    foreach (DataTable table in dstable.Tables)
                    {
                        var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                        var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                        sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                        DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                        string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                        if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                        {
                            sheetId =
                                sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                        }

                        DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                        sheets.Append(sheet);

                        DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                        List<String> columns = new List<string>();
                        foreach (DataColumn column in table.Columns)
                        {
                            columns.Add(column.ColumnName);

                            DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                            cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName);
                            headerRow.AppendChild(cell);
                        }

                        sheetData.AppendChild(headerRow);

                        foreach (DataRow dsrow in table.Rows)
                        {
                            DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                            foreach (String col in columns)
                            {
                                DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()); //
                                newRow.AppendChild(cell);
                            }

                            sheetData.AppendChild(newRow);
                        }

                    }
                }

                status = true;
            }
            catch (Exception ex)
            {

            }

            return status;
        }
        public static string DatatableToExcel(DataTable table, int exportSheetType = 0, bool allLeftAlign = false, bool isErrorOutput = false)
        {
            try
            {

                StringBuilder strHTMLBuilder = new StringBuilder();
                strHTMLBuilder.Append("<html >");
                strHTMLBuilder.Append("<head>");
                if (allLeftAlign)
                {
                    strHTMLBuilder.Append("<Style> table,th, tr,td {text-align: left; } </Style>");
                }
                strHTMLBuilder.Append("</head>");
                strHTMLBuilder.Append("<body>");
                strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='1' style='font-family:Verdana; font-size:smaller;border-collapse:collapse;'>");

                strHTMLBuilder.Append("<tr style='font-weight:bold; height:30px; font-size:normal'>");
                foreach (DataColumn myColumn in table.Columns)
                {
                    if (exportSheetType.Equals(7) && myColumn.ColumnName.ToLower().Contains("note")) //trip dupicate
                    {
                        strHTMLBuilder.Append("<td style='background-color:#a0c4ff; width:550px'>");
                    }
                    else
                    {
                        if (isErrorOutput)
                            strHTMLBuilder.Append("<td style='background-color:#f94367'>");
                        else
                            strHTMLBuilder.Append("<td style='background-color:#a0c4ff;'>");
                    }
                    strHTMLBuilder.Append(myColumn.ColumnName);
                    strHTMLBuilder.Append("</td>");

                }
                strHTMLBuilder.Append("</tr>");

                foreach (DataRow myRow in table.Rows)
                {
                    strHTMLBuilder.Append("<tr>");
                    foreach (DataColumn myColumn in table.Columns)
                    {
                        if (myRow[myColumn.ColumnName].ToString().ToLower().Contains("sheeowtal")) //for DTC frequency generator sheet
                        {
                            strHTMLBuilder.Append("<td style='font-weight:bold;height:35px; text-align:center; font-size:normal; background:#fffed3' colspan='" + table.Columns.Count + "'>");
                            strHTMLBuilder.Append(myRow[myColumn.ColumnName].ToString());
                            strHTMLBuilder.Append("</td>");
                            break;
                        }
                        else if (myColumn.ColumnName.ToLower().Contains("bold")) // for DTC frequency generator sheet
                        {
                            strHTMLBuilder.Append("<td style='font-weight:bold;'>");
                            strHTMLBuilder.Append(myRow[myColumn.ColumnName].ToString());
                            strHTMLBuilder.Append("</td>");
                        }
                        else if ((exportSheetType.Equals(1) || exportSheetType.Equals(3)) && myRow[myColumn.ColumnName].ToString().ToLower().Equals("1")) // for find distance sheet
                        {
                            strHTMLBuilder.Append("<td style='font-weight:bold; font-size:large; background:#bffcab'>");
                            strHTMLBuilder.Append(myRow[myColumn.ColumnName].ToString());
                            strHTMLBuilder.Append("</td>");
                        }
                        else if (myRow[myColumn.ColumnName].ToString().ToLower().Contains("^file")) //for BMTC-Bangalore
                        {
                            strHTMLBuilder.Append("<td style='font-weight:bold;height:35px; text-align:center; font-size:normal; background:#fffed3' colspan='" + table.Columns.Count + "'>");
                            strHTMLBuilder.Append(myRow[myColumn.ColumnName].ToString().Replace("^", ""));
                            strHTMLBuilder.Append("</td>");
                            break;
                        }
                        else if (myRow[myColumn.ColumnName].ToString().ToLower().Contains("error")) //for BMTC-Bangalore
                        {
                            strHTMLBuilder.Append("<td style='height:50px; text-align:center; font-size:normal;   color:#6d0000' colspan='" + table.Columns.Count + "'>");
                            strHTMLBuilder.Append(myRow[myColumn.ColumnName].ToString());
                            strHTMLBuilder.Append("</td>");
                            break;
                        }
                        else if (exportSheetType.Equals(7) && myRow[myColumn.ColumnName].ToString().ToLower().Contains("summery")) //for trip dupicate
                        {
                            strHTMLBuilder.Append("<td style='height:70px; text-align:center; font-size:15px;   color:#6d0000' colspan='" + table.Columns.Count + "'>");
                            strHTMLBuilder.Append(myRow[myColumn.ColumnName].ToString());
                            strHTMLBuilder.Append("</td>");
                            break;
                        }

                        else if (exportSheetType.Equals(8) && myRow[myColumn.ColumnName].ToString().ToLower().Contains("red")) //for route track creation
                        {
                            strHTMLBuilder.Append("<td style='background-color:#ffbfbf;'>");
                            strHTMLBuilder.Append(myRow[myColumn.ColumnName].ToString().Replace("Red", ""));
                            strHTMLBuilder.Append("</td>");
                        }
                        else if (exportSheetType.Equals(9) && myRow[myColumn.ColumnName].ToString().ToLower().Contains("origin")) //for fare matrix
                        {
                            strHTMLBuilder.Append("<td style='background-color:#00FF00'>");
                            strHTMLBuilder.Append(myRow[myColumn.ColumnName].ToString().ToLower().Replace("origin", ""));
                            strHTMLBuilder.Append("</td>");
                        }

                        else if (exportSheetType.Equals(10) && myColumn.ColumnName.ToLower().Equals("stop_seq") && myRow[myColumn.ColumnName].ToString().Equals("1")) //for KSRTC Mysore Schedule
                        {
                            strHTMLBuilder.Append("<td style='background-color:#d0f9a4'>");
                            strHTMLBuilder.Append("1");
                            strHTMLBuilder.Append("</td>");
                        }
                        else if (exportSheetType.Equals(10) && (myRow[myColumn.ColumnName].ToString().ToLower().Equals("schedulemismatched") || myRow[myColumn.ColumnName].ToString().ToLower().Equals("duplicatetime"))) //for KSRTC Mysore Schedule
                        {
                            strHTMLBuilder.Append("<td style='background-color:#f95979'>");
                            strHTMLBuilder.Append(myRow[myColumn.ColumnName].ToString());
                            strHTMLBuilder.Append("</td>");
                        }
                        else
                        {
                            strHTMLBuilder.Append("<td>");
                            strHTMLBuilder.Append(myRow[myColumn.ColumnName].ToString());
                            strHTMLBuilder.Append("</td>");
                        }
                    }
                    strHTMLBuilder.Append("</tr>");
                }

                //Close tags.  
                strHTMLBuilder.Append("</table>");
                strHTMLBuilder.Append("</body>");
                strHTMLBuilder.Append("</html>");

                return strHTMLBuilder.ToString();
            }
            catch
            {
                return string.Empty;
            }


        }
        public static DateTime UnixTimeStampToLocalDateTime(string unixTimeStamp)
        {
            // Unix timestamp is seconds past epoch
            System.DateTime dtDateTime = new DateTime(1970, 1, 1, 0, 0, 0, 0, System.DateTimeKind.Utc);
            dtDateTime = dtDateTime.AddSeconds(Convert.ToDouble(unixTimeStamp)).ToLocalTime();
            return dtDateTime;
        }
    }
}
