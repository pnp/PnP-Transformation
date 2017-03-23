using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Collections;
using Microsoft.Office.Interop.Excel;
using System.Xml;
using System.Runtime.InteropServices;

using Microsoft.SharePoint.Client;
using JDP.Remediation.Console.Common.CSV;
using JDP.Remediation.Console.Common.Base;
using JDP.Remediation.Console.Common.Utilities;

namespace JDP.Remediation.Console.PivotHelper
{
    class SummaryViewHelper
    {
        /// <summary>
        /// Creates SummaryView sheet
        /// </summary>
        /// <param name="summaryViewNode"></param>
        /// <param name="dSummaryViewComponents"></param>
        /// <param name="PivotOutputReportFullPath"></param>
        public static void ComponentsSummaryView(XmlNode summaryViewNode, Dictionary<string, string[]> dSummaryViewComponents, string PivotOutputReportFullPath)
        {
            Excel.Application oApp;
            Excel.Workbook componentIpBook = null;
            Excel.Worksheet componentSheet = null; ;
            Excel.Worksheet summaryViewSheet = null;
            Excel.Workbook summaryBook = null;
            Excel.Range range = null;

            List<string> lstAllWebAppUrls = new List<string>();
            List<List<string>> lstAllCompWebAppUrls = new List<List<string>>();

            //SummaryView sheet attributes
            int row = 0;
            int col = 0;
            char cellIndex = 'A';
            string grandTotalLabel = string.Empty;
            string summaryViewPreMtColumn = string.Empty;
            int fontSize = 10;
            string fontName = string.Empty;
            string fontColor = string.Empty;
            double rowHeight = 12;

            //SummaryViewColumn,SummaryViewColumnHeader attribute values
            string summaryViewColumnText = string.Empty;
            string summaryViewHeadingText = string.Empty;
            string summaryViewHeadingCell = string.Empty;
            string summaryViewDescriptionStartCell = string.Empty;
            string summaryViewDescriptionEndCell = string.Empty;
            int summaryViewHeaderFontSize = 10;
            string summaryViewHeaderFontColor = string.Empty;
            int summaryViewColumnHeaderFontSize = 10;
            string summaryViewColumnHeaderFontColor = string.Empty;
            int summaryViewHeadingRow = 0;
            int summaryViewHeadingColumn = 0;
            string summaryViewHeadingBgColor = string.Empty;
            string summaryViewColumnBgColor = string.Empty;
            string summaryViewDescriptionText = string.Empty;

            //Chart Attributes
            int chartStyle = 0;
            int chartLayoutType = 0;
            int chartColor = 0;
            int chartWidth = 0;
            int chartHeight = 0;
            string chartType = string.Empty;

            try
            {
                if (summaryViewNode.HasChildNodes)
                {
                    row = Convert.ToInt32(summaryViewNode.SelectSingleNode("StartingRow").InnerText);
                    col = Convert.ToInt32(summaryViewNode.SelectSingleNode("StartingColumn").InnerText);
                    cellIndex = Convert.ToChar(summaryViewNode.SelectSingleNode("CellIndex").InnerText);
                    grandTotalLabel = summaryViewNode.SelectSingleNode("GrandTotalLabel").InnerText;
                    fontSize = Convert.ToInt16(summaryViewNode.SelectSingleNode("FontSize").InnerText);
                    fontName = summaryViewNode.SelectSingleNode("FontName").InnerText;
                    fontColor = summaryViewNode.SelectSingleNode("FontColor").InnerText;
                    summaryViewDescriptionText = summaryViewNode.SelectSingleNode("Description").InnerText;
                    rowHeight = Convert.ToDouble(summaryViewNode.SelectSingleNode("SummaryViewHeading").Attributes["RowHeight"].InnerText);

                    if (summaryViewNode.SelectSingleNode("SummaryViewColumnHeader").Attributes.Count > 0)
                    {
                        summaryViewColumnText = summaryViewNode.SelectSingleNode("SummaryViewColumnHeader").Attributes["Text"].InnerText;
                        summaryViewColumnBgColor = summaryViewNode.SelectSingleNode("SummaryViewColumnHeader").Attributes["BgColor"].InnerText;
                        summaryViewColumnHeaderFontSize = Convert.ToInt16(summaryViewNode.SelectSingleNode("SummaryViewColumnHeader").Attributes["FontSize"].InnerText);
                        summaryViewColumnHeaderFontColor = summaryViewNode.SelectSingleNode("SummaryViewColumnHeader").Attributes["FontColor"].InnerText;
                    }
                    if (summaryViewNode.SelectSingleNode("SummaryViewHeading").Attributes.Count > 0)
                    {
                        summaryViewHeadingText = summaryViewNode.SelectSingleNode("SummaryViewHeading").Attributes["Text"].InnerText;
                        summaryViewHeadingCell = summaryViewNode.SelectSingleNode("SummaryViewHeading").Attributes["Cell"].InnerText;
                        summaryViewDescriptionStartCell = summaryViewNode.SelectSingleNode("SummaryViewHeading").Attributes["DescriptionStartCell"].InnerText;
                        summaryViewDescriptionEndCell = summaryViewNode.SelectSingleNode("SummaryViewHeading").Attributes["DescriptionEndCell"].InnerText;
                        summaryViewHeadingRow = Convert.ToInt32(summaryViewNode.SelectSingleNode("SummaryViewHeading").Attributes["Row"].InnerText);
                        summaryViewHeadingColumn = Convert.ToInt32(summaryViewNode.SelectSingleNode("SummaryViewHeading").Attributes["Column"].InnerText);
                        summaryViewHeadingBgColor = summaryViewNode.SelectSingleNode("SummaryViewHeading").Attributes["BgColor"].InnerText;
                        summaryViewHeaderFontSize = Convert.ToInt16(summaryViewNode.SelectSingleNode("SummaryViewHeading").Attributes["FontSize"].InnerText);
                        summaryViewHeaderFontColor = summaryViewNode.SelectSingleNode("SummaryViewHeading").Attributes["FontColor"].InnerText;
                    }
                    if (summaryViewNode.SelectSingleNode("Chart").Attributes.Count > 0)
                    {
                        chartStyle = Convert.ToInt32(summaryViewNode.SelectSingleNode("Chart").Attributes["Style"].InnerText);
                        chartLayoutType = Convert.ToInt32(summaryViewNode.SelectSingleNode("Chart").Attributes["LayoutType"].InnerText);
                        chartColor = Convert.ToInt32(summaryViewNode.SelectSingleNode("Chart").Attributes["Color"].InnerText);
                        chartWidth = Convert.ToInt32(summaryViewNode.SelectSingleNode("Chart").Attributes["Width"].InnerText);
                        chartHeight = Convert.ToInt32(summaryViewNode.SelectSingleNode("Chart").Attributes["Height"].InnerText);
                        XmlNodeList xnList = summaryViewNode.SelectSingleNode("Chart").SelectNodes("ChartType");
                        foreach (XmlNode node in xnList)
                        {
                            XmlElement chartTypeElement = node as XmlElement;
                            if (chartTypeElement.HasAttribute("active"))
                            {
                                chartType = chartTypeElement.InnerText;
                                break;
                            }
                        }

                        if (string.IsNullOrEmpty(chartType))
                            chartType = "3dpie";
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(string.Format("[GeneratePivotReports][ComponentsSummaryView: Reading xml values][Exception]: " + ex.Message + "\n" + ex.StackTrace.ToString()), true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Pivot", ex.Message, ex.ToString(),
                             "[GeneratePivotReports]: ComponentsSummaryView", ex.GetType().ToString(), "Reading xml values");
            }

            oApp = new Excel.Application();
            var workbooks = oApp.Workbooks;

            Logger.LogInfoMessage(string.Format("[GeneratePivotReports][ComponentsSummaryView] Processing Started to add details in Summary View Sheet"), false);

            try
            {
                if (System.IO.File.Exists(PivotOutputReportFullPath))
                {
                    summaryBook = oApp.Workbooks.Open(PivotOutputReportFullPath);
                    // Create multiple sheets
                    if (oApp.Application.Application.Sheets.Count >= 1)
                    {
                        summaryViewSheet = (Excel.Worksheet)summaryBook.Worksheets.Add(Missing.Value, summaryBook.Worksheets[1], Missing.Value, Missing.Value);
                        summaryViewSheet.Name = summaryViewNode.Name;
                        //Constants.SummaryViewSheetName;
                    }
                    else
                    {
                        summaryViewSheet = oApp.Worksheets[2];
                    }
                }


                CreateHeaders(row, col, summaryViewColumnText, cellIndex + row.ToString(), summaryViewColumnBgColor, true,
                    summaryViewColumnHeaderFontSize, summaryViewColumnHeaderFontColor, fontName, summaryViewSheet);
                int colCount = 0;

                //Get list of component names and write as header for SummaryView
                foreach (var component in dSummaryViewComponents)
                {
                    string componentName = component.Key.ToString();
                    componentName = componentName.Substring(componentName.IndexOf("~") + 1);
                    string componentFileName = component.Value[0].ToString();
                    List<string> lstComponentWebAppUrls = null;
                    colCount++;
                    if (System.IO.File.Exists(componentFileName))
                    {
                        componentIpBook = oApp.Workbooks.Open(componentFileName);
                        componentSheet = (Excel.Worksheet)componentIpBook.Worksheets[1];
                        range = componentSheet.UsedRange;

                        lstComponentWebAppUrls = GetWebApplicationUrls(range, component.Value[1].ToString());

                        lstAllCompWebAppUrls.Add(lstComponentWebAppUrls);
                        if (lstComponentWebAppUrls != null)
                        {
                            if (lstComponentWebAppUrls.Count > 0)
                                lstAllWebAppUrls.AddRange(lstComponentWebAppUrls);
                        }
                        //FillComponentHeader(ref oOutputSheet, Constants.WebApplicationColumn + startingRow, "WebApplicationUrls");
                        WriteHeaders(componentName, row, col + colCount, (Char)(Convert.ToUInt16(cellIndex) + colCount),
                            summaryViewColumnBgColor, summaryViewColumnHeaderFontSize, summaryViewColumnHeaderFontColor,
                            fontName, summaryViewSheet);
                        componentIpBook.Close(false, Missing.Value, Missing.Value);
                        Marshal.ReleaseComObject(componentIpBook);
                    }
                }
                //Get unique WebApplicationUrls
                lstAllWebAppUrls = lstAllWebAppUrls.Distinct().ToList();
                int noOfWebUrls = 0;
                //Get count of unique WebApplicationUrls
                if (lstAllWebAppUrls != null)
                    noOfWebUrls = lstAllWebAppUrls.Count;
                int rowCount = row;
                colCount = 0;

                foreach (var item in lstAllWebAppUrls)
                {
                    rowCount++;
                    //Get the cell name to fill WebApplicationUrl
                    string webUrlFillCell = ((Char)(Convert.ToUInt16(cellIndex))).ToString() + rowCount.ToString();
                    var fillRange = summaryViewSheet.get_Range(webUrlFillCell, webUrlFillCell);
                    fillRange.Font.Size = fontSize;
                    fillRange.Font.Name = fontName;
                    fillRange.Value2 = item.ToString();

                    foreach (List<string> componentWebUrls in lstAllCompWebAppUrls)
                    {
                        colCount++;
                        //Get the cell name to fill WebApplicationUrl count
                        string cellName = ((Char)(Convert.ToUInt16(cellIndex) + colCount)).ToString() + rowCount.ToString();
                        //Fill WebApplicationUrl count
                        FillWebApplicationCount(componentWebUrls, summaryViewSheet, item.ToString(), cellName, fontSize, fontName);

                    }
                    colCount = 0;
                }

                //Fill GrandTotal of WebApplicationUrls of each component
                FillGrandTotal(summaryViewSheet, ((Char)(Convert.ToUInt16(cellIndex))).ToString() + (row + noOfWebUrls + 1).ToString(),
                    grandTotalLabel, fontSize, fontName, summaryViewColumnBgColor, summaryViewColumnHeaderFontColor);

                foreach (List<string> componentWebUrls in lstAllCompWebAppUrls)
                {
                    colCount++;
                    if (componentWebUrls != null)
                        FillGrandTotal(summaryViewSheet, ((Char)(Convert.ToUInt16(cellIndex) + colCount)).ToString() + (row + noOfWebUrls + 1).ToString(),
                            componentWebUrls.Count.ToString(), fontSize, fontName, summaryViewColumnBgColor, summaryViewColumnHeaderFontColor);
                    else
                        FillGrandTotal(summaryViewSheet, ((Char)(Convert.ToUInt16(cellIndex) + colCount)).ToString() + (row + noOfWebUrls + 1).ToString(), "0",
                            fontSize, fontName, summaryViewColumnBgColor, summaryViewColumnHeaderFontColor);
                }
                colCount = 0;

                //Create Graph for each component
                foreach (var component in dSummaryViewComponents)
                {
                    colCount++;
                    ChartHelper.DrawGraph(component.Key.ToString(), summaryViewSheet, cellIndex.ToString(), colCount,
                        row + noOfWebUrls, row, chartType.ToLower(), chartWidth, chartHeight, chartStyle, cellIndex);
                }

                //Write Summary View Header
                CreateHeaders(summaryViewHeadingRow, summaryViewHeadingColumn, summaryViewHeadingText, summaryViewHeadingCell,
                    summaryViewHeadingBgColor, true, summaryViewHeaderFontSize, summaryViewHeaderFontColor, fontName, summaryViewSheet);

                //Write Summary View Description
                Excel.Range summaryViewDescriptionRange = summaryViewSheet.get_Range(summaryViewDescriptionStartCell, summaryViewDescriptionEndCell);
                summaryViewDescriptionRange.Merge();
                summaryViewDescriptionRange.Value = summaryViewDescriptionText.Trim();
                summaryViewDescriptionRange.Columns.AutoFit();
                summaryViewDescriptionRange.RowHeight = rowHeight;
                summaryViewDescriptionRange.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
                summaryViewDescriptionRange.WrapText = true;
                summaryViewDescriptionRange.Font.Size = Convert.ToInt16(fontSize);
                summaryViewDescriptionRange.Font.Name = fontName;
                summaryViewDescriptionRange.Borders.Color = XlRgbColor.rgbBlack;
                summaryViewDescriptionRange.Font.Color = CommonHelper.GetColor(fontColor);

                //Autofit first column of the sheet
                Excel.Range er = summaryViewSheet.get_Range("A:A", System.Type.Missing);
                er.EntireColumn.AutoFit();

                //Release Object/ Dispose Objects
                summaryBook.Save();
                summaryBook.Close();
                workbooks.Close();
                oApp.Application.Quit();
                oApp.Quit();

                Marshal.ReleaseComObject(componentSheet);
                Marshal.ReleaseComObject(summaryViewSheet);

                Marshal.ReleaseComObject(summaryBook);
                Marshal.ReleaseComObject(workbooks);

                Logger.LogInfoMessage(string.Format("[GeneratePivotReports][ComponentsSummaryView] Process Completed to add details in Summary View Sheet"), true);
            }
            catch (Exception ex)
            {
                if (summaryBook != null)
                {
                    summaryBook.Save();
                    summaryBook.Close();
                }

                object misValue = System.Reflection.Missing.Value;
                if (componentIpBook != null)
                {
                    componentIpBook.Close(false, misValue, misValue);
                }
                if (oApp != null)
                {
                    oApp.Quit();
                    oApp.Application.Quit();
                }

                Marshal.ReleaseComObject(componentIpBook);
                Marshal.ReleaseComObject(workbooks);
                Marshal.ReleaseComObject(summaryBook);

                Logger.LogErrorMessage(string.Format("[GeneratePivotReports][ComponentsSummaryView][Exception]: " + ex.Message + "\n" + ex.StackTrace.ToString()), true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Pivot", ex.Message, ex.ToString(),
                            "[GeneratePivotReports]: ComponentsSummaryView", ex.GetType().ToString(), Constants.NotApplicable);
            }
            finally
            {
                Marshal.ReleaseComObject(oApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();

                lstAllCompWebAppUrls = null;
                lstAllWebAppUrls = null;
                range = null;
                componentSheet = null;
                componentIpBook = null;
                summaryViewSheet = null;
                summaryBook = null;
                oApp = null;
            }
        }


        /// <summary>
        /// Returns list of WebApplicationUrls
        /// </summary>
        /// <param name="range">Used range of component input file sheet</param>
        /// <param name="summaryViewColumnText"></param>
        /// <returns></returns>
        public static List<string> GetWebApplicationUrls(Excel.Range range, string summaryViewColumnText)
        {
            List<string> lstWebApplicationUrls = null;
            try
            {
                if (range.Rows.Count > 1)
                {
                    for (int i = 1; i <= range.Columns.Count; i++)
                    {
                        try
                        {
                            string str = (string)(range.Cells[1, i] as Excel.Range).Value2;
                            //Find for WebApplicationUrl column name
                            if (str.Equals(summaryViewColumnText))
                            {
                                Excel.Range ctRange2 = range.Columns[i];
                                //Get all WebApplicationUrl values into array
                                System.Array ctWebAppValue = (System.Array)ctRange2.Cells.Value;
                                //Convert array into List<string>
                                lstWebApplicationUrls = ctWebAppValue.OfType<object>().Select(o => o.ToString()).ToList();
                                //Remove header name 'WebApplicationUrl'
                                lstWebApplicationUrls.Remove(lstWebApplicationUrls.First());
                                List<string> lstNewWebAppUrls = new List<string>();
                                //Trim '/' character when WebApplication ends with '/'
                                lstNewWebAppUrls = lstWebApplicationUrls.Select(item => item.EndsWith("/") ? item.TrimEnd('/') : item).ToList();

                                return lstNewWebAppUrls;
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.LogErrorMessage(string.Format("[GeneratePivotReports][ComponentsSummaryView][GetWebApplicationUrls][Exception]: " + ex.Message + "\n" + ex.StackTrace.ToString()), true);
                            ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Pivot", ex.Message, ex.ToString(),
                            "[GeneratePivotReports]: GetWebApplicationUrls", ex.GetType().ToString(), Constants.NotApplicable);
                        }
                    }
                }
                else
                    return null;
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(string.Format("[GeneratePivotReports][ComponentsSummaryView][GetWebApplicationUrls][Exception]: " + ex.Message + "\n" + ex.StackTrace.ToString()), true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Pivot", ex.Message, ex.ToString(),
                            "[GeneratePivotReports]: GetWebApplicationUrls", ex.GetType().ToString(), Constants.NotApplicable);
            }
            finally
            {
                range = null;
            }
            return lstWebApplicationUrls;
        }

        /// <summary>
        /// Fills WebApplicationUrl count of component for given webApplicaitonUrl value
        /// </summary>
        /// <param name="lstComponentWebAppUrls">Contains list of WebApplicationUrls of the component</param>
        /// <param name="oSummaryViewSheet">sheet object in which the count to fill</param>
        /// <param name="webAppUrl">WebApplicationUrl value for the count</param>
        /// <param name="componentColumn">Column name/index where to fill the count value</param>
        /// <param name="fontSize"></param>
        /// <param name="fontName"></param>
        public static void FillWebApplicationCount(List<string> lstComponentWebAppUrls, Excel.Worksheet oSummaryViewSheet,
            string webAppUrl, string componentColumn, int fontSize, string fontName)
        {
            try
            {
                //Get the range of sheet to fill count
                var ctWebAppCountRange = oSummaryViewSheet.get_Range(componentColumn, componentColumn);
                if (lstComponentWebAppUrls != null)
                {
                    //Check whether list containing given WebApplicationUrl, if not fill with value 0
                    if (lstComponentWebAppUrls.Contains(webAppUrl, StringComparer.CurrentCultureIgnoreCase))
                    {
                        //Get the count of WebApplicationUrl
                        int curWebAppCTCount = lstComponentWebAppUrls.Where(o => o.ToString().Equals(webAppUrl.ToString())).Count();

                        ctWebAppCountRange.Value2 = curWebAppCTCount.ToString();
                    }
                    else
                    {
                        ctWebAppCountRange.Value2 = "0";
                    }
                }
                else
                {
                    ctWebAppCountRange.Value2 = "0";
                }
                ctWebAppCountRange.Font.Size = fontSize;
                ctWebAppCountRange.Font.Name = fontName;
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage("[GeneratePivotReports][ComponentsSummaryView][FillWebApplicationCount][Exception]: " + ex.Message + "\n" + ex.StackTrace.ToString(), true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Pivot", ex.Message, ex.ToString(),
                            "[GeneratePivotReports]: FillWebApplicationCount", ex.GetType().ToString(), Constants.NotApplicable);
            }
            finally
            {
                lstComponentWebAppUrls = null;
                oSummaryViewSheet = null;
            }
        }

        /// <summary>
        /// Fills GrandTotal of WebApplicationUrls of given component column
        /// </summary>
        /// <param name="oSummarySheet"></param>
        /// <param name="componentColumn"></param>
        /// <param name="grandTotal"></param>
        /// <param name="fontSize"></param>
        /// <param name="fontName"></param>
        /// <param name="bgColor"></param>
        /// <param name="fontColor"></param>
        public static void FillGrandTotal(Excel.Worksheet oSummarySheet, string componentColumn, string grandTotal,
            int fontSize, string fontName, string bgColor, string fontColor)
        {
            try
            {
                //Get the Range where to fill total value
                var grandTotalRange = oSummarySheet.get_Range(componentColumn, componentColumn);
                grandTotalRange.Interior.Color = CommonHelper.GetColor(bgColor);
                grandTotalRange.Borders.Color = XlRgbColor.rgbBlack;
                grandTotalRange.Font.Bold = true;
                grandTotalRange.Font.Size = fontSize;
                grandTotalRange.Font.Name = fontName;
                grandTotalRange.Font.Color = CommonHelper.GetColor(fontColor);
                grandTotalRange.Value2 = grandTotal;
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(string.Format("[GeneratePivotReports][ComponentsSummaryView][FillGrandTotal][Exception]: " + ex.Message + "\n" + ex.StackTrace.ToString()), true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Pivot", ex.Message, ex.ToString(),
                            "[GeneratePivotReports]: FillGrandTotal", ex.GetType().ToString(), Constants.NotApplicable);

            }
            finally
            {
                oSummarySheet = null;
            }
        }

        /// <summary>
        /// Writes headers of component
        /// </summary>
        /// <param name="component"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="cell"></param>
        /// <param name="oSummarySheet"></param>
        public static void WriteHeaders(string component, int row, int col, char cell, string bgColor,
            int fontSize, string fontColor, string fontName, Excel.Worksheet oSummarySheet)
        {
            CreateHeaders(row, col, component, cell + row.ToString(), bgColor, true, fontSize, fontColor, fontName, oSummarySheet);
            oSummarySheet = null;
        }

        /// <summary>
        /// Creates header for component
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="htext"></param>
        /// <param name="cell"></param>
        /// <param name="mergeColumns"></param>
        /// <param name="backgroundColor"></param>
        /// <param name="bold"></param>
        /// <param name="fontSize"></param>
        /// <param name="fcolor"></param>
        /// <param name="oSummarySheet"></param>
        public static void CreateHeaders(int row, int col, string htext, string cell,
            string backgroundColor, bool bold, int fontSize, string fcolor, string fontName, Excel.Worksheet oSummarySheet)
        {
            try
            {
                oSummarySheet.Cells[row, col] = htext;
                Excel.Range workSheet_range = null;
                workSheet_range = oSummarySheet.get_Range(cell, cell);
                workSheet_range.Borders.Color = XlRgbColor.rgbBlack;
                workSheet_range.Font.Bold = bold;
                workSheet_range.Font.Size = fontSize;
                workSheet_range.Font.Name = fontName;
                //set background color
                workSheet_range.Interior.Color = CommonHelper.GetColor(backgroundColor);
                //set font color
                workSheet_range.Font.Color = CommonHelper.GetColor(fcolor);
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(string.Format("[GeneratePivotReports][ComponentsSummaryView][CreateHeaders][Exception]: " + ex.Message + "\n" + ex.StackTrace.ToString()), true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Pivot", ex.Message, ex.ToString(),
                            "[GeneratePivotReports]: CreateHeaders", ex.GetType().ToString(), Constants.NotApplicable);
            }
            finally
            {
                oSummarySheet = null;
            }
        }
    }
}
