using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
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
    class PivotViewHelper
    {
        /// <summary>
        /// Creates all the pivot tables, slicers, and row fields, filters and Description, based on values from 
        /// their hashTables.
        /// </summary>
        /// <param name="inputCSVFile"></param>
        /// <param name="PivotOutputReportFullPath"></param>
        /// <param name="newSheetName"></param>
        /// <param name="inputDataSheetName"></param>
        /// <param name="PivottableName"></param>
        /// <param name="htRowPivotFields"></param>
        /// <param name="htPgaefilterFields"></param>
        /// <param name="htSlicers"></param>
        /// <param name="pivotCountField"></param>
        /// <param name="component"></param>
        /// <param name="numberOfFilesCount"></param>
        /// <param name="otherNodes"></param>
        public static void GeneratePivotAndSlicersView(string inputCSVFile, string PivotOutputReportFullPath, ref string newSheetName, string inputDataSheetName, string PivottableName, Hashtable htRowPivotFields, Hashtable htPgaefilterFields, Hashtable htSlicers, string pivotCountField, XmlNode component, int numberOfFilesCount, List<XmlNode> otherNodes)
        {
            Excel.Application oApp;
            Excel.Worksheet oSheet;
            Excel.Workbook oBook = null;
            Excel.Worksheet oSheet1 = null;
            Excel.Worksheet oSheet2 = null;
            oApp = new Excel.Application();
            var workbooks = oApp.Workbooks;
            string sheetName = newSheetName;
            XmlNode style = otherNodes.Find(item => item.Name == "Style");
            Excel.Range rangeToChange = null;

            string exceptionComment = "Processing for CSV :" + inputCSVFile;
            Logger.LogInfoMessage(string.Format("[GeneratePivotReports][GeneratePivotAndSlicersView] Processing Started for (" + inputCSVFile + ")"), false);
            Excel.Workbook oDiscoveryViewook = null;
            try
            {
                oBook = oApp.Workbooks.Open(inputCSVFile);
                //Excel.Workbook oDiscoveryViewook = null;

                if (System.IO.File.Exists(PivotOutputReportFullPath))
                {
                    oDiscoveryViewook = workbooks.Open(PivotOutputReportFullPath);
                    // create multiple sheets
                    if (oApp.Application.Application.Sheets.Count >= 1)
                    {
                        oSheet2 = oDiscoveryViewook.Worksheets.OfType<Excel.Worksheet>().FirstOrDefault(ws => ws.Name == sheetName);
                        oSheet1 = (Excel.Worksheet)oDiscoveryViewook.Worksheets.Add(After: oDiscoveryViewook.Sheets[oDiscoveryViewook.Sheets.Count]);
                        if (oSheet2 == null)
                        {
                            oSheet1.Name = newSheetName;
                        }
                        else
                        {
                            if (newSheetName.Length >= Constants.SheetNameMaxLength)
                                newSheetName = newSheetName.Remove(newSheetName.Length - 3);

                            newSheetName = newSheetName + "_" + numberOfFilesCount;
                            oSheet1.Name = newSheetName;
                        }
                    }
                    else
                    {
                        oSheet1 = oApp.Worksheets[2];
                    }
                }

                if (inputDataSheetName.Length >= Constants.SheetNameMaxLength)
                    oSheet = (Excel.Worksheet)oBook.Sheets.get_Item(1);
                else
                    oSheet = (Excel.Worksheet)oBook.Sheets.get_Item(inputDataSheetName);


                // now capture range of the first sheet 
                Excel.Range oRange = oSheet.UsedRange;
                // specify first cell for pivot table       
                Excel.Range oRange2 = (Excel.Range)oSheet1.Cells[3, 1];
                //Create Pivot Cache

                if (oRange.Rows.Count > 1)
                {
                    PivotCache oPivotCache = oDiscoveryViewook.PivotCaches().Create(XlPivotTableSourceType.xlDatabase, oRange, XlPivotTableVersionList.xlPivotTableVersion14);
                    PivotTable oPivotTable = oPivotCache.CreatePivotTable(TableDestination: oRange2, TableName: PivottableName);

                    //Creating row pivot fields
                    foreach (DictionaryEntry rowField in htRowPivotFields)
                    {
                        string rowFieldContent = rowField.Value.ToString();
                        string rowFieldValue = rowFieldContent.Substring(0, rowFieldContent.IndexOf("~"));
                        string rowFieldLabel = rowFieldContent.Substring(rowFieldContent.IndexOf("~") + 1);
                        if ((Excel.PivotField)oPivotTable.PivotFields(Convert.ToString(rowFieldValue)) != null)
                        {
                            Excel.PivotField oPivotFieldPivotFieldName = (Excel.PivotField)oPivotTable.PivotFields(Convert.ToString(rowFieldValue));
                            oPivotFieldPivotFieldName.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                            oPivotTable.CompactLayoutRowHeader = rowFieldLabel;
                        }
                        else
                            Logger.LogInfoMessage(string.Format("[GeneratePivotReports][GeneratePivotAndSlicersView] Error: No data available"), true);
                    }

                    //page filters
                    foreach (DictionaryEntry rowPgaeField in htPgaefilterFields)
                    {
                        Excel.PivotField scPageFilterFiled = oPivotTable.PivotFields(Convert.ToString(rowPgaeField.Value));
                        scPageFilterFiled.Orientation = Excel.XlPivotFieldOrientation.xlPageField;
                    }

                    //Count Field
                    Excel.PivotField oPivotField2 = (Excel.PivotField)oPivotTable.PivotFields(pivotCountField);
                    oPivotTable.AddDataField(oPivotField2, "Count of " + pivotCountField + "", Excel.XlConsolidationFunction.xlCount);

                    rangeToChange = oPivotTable.TableRange2;

                    oSheet1.Activate();
                    rangeToChange.Select();
                    //excelApp.Selection.Font;
                    oSheet1.Application.Selection.Font.Name = CommonHelper.CheckAttributes("FontFamily", null, style);
                    oSheet1.Application.Selection.Font.Size = Convert.ToDouble(CommonHelper.CheckAttributes("FontSize", null, style));

                    //Create Slicer Cache Object    
                    int Slicerpos = 0, slicersCount = 0;
                    foreach (DictionaryEntry rowSlicer in htSlicers)
                    {
                        slicersCount++;
                        string rowSlicerValue = rowSlicer.Value.ToString();
                        if (rowSlicerValue.Contains("~"))
                        {
                            string sliceName = rowSlicerValue.Substring(0, rowSlicerValue.IndexOf("~"));
                            string sliceStyle = rowSlicerValue.Substring(rowSlicerValue.IndexOf("~") + 1);
                            Excel.SlicerCache oSlicerCache = (Excel.SlicerCache)oDiscoveryViewook.SlicerCaches.Add2(oPivotTable, sliceName);
                            Excel.Slicer oSlicer = (Excel.Slicer)oSlicerCache.Slicers.Add(oSheet1, Type.Missing, sliceName + "_" + newSheetName, sliceName, Top: 30, Left: 400 + Slicerpos, Width: 144, Height: 200);
                            oSlicer.Style = sliceStyle;

                            //To Move Left Position of next slicers(2,3...)
                            Slicerpos += 190;
                        }
                    }
                }
                Range Line = (Range)oSheet1.Rows[1];
                Line.Insert();

                XmlNode descriptionTitle = component.SelectSingleNode("DescriptionTitle");
                XmlNode description = component.SelectSingleNode("Description");

                //Get the range of sheet to fill count
                XmlNode styleOfDescription = style.SelectSingleNode("ComponentStyle").SelectSingleNode("Description");

                Excel.Range descriptionRange = oSheet1.get_Range("B1", "O1");
                descriptionRange.Merge();

                string descText = description.InnerText.Trim();
                descriptionRange.Value = descText.Replace("\r", "").Replace("\n", "");
                descriptionRange.Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
                descriptionRange.Columns.AutoFit();
                descriptionRange.RowHeight = Convert.ToDouble(CommonHelper.CheckAttributes("RowHeight", styleOfDescription, style));
                descriptionRange.WrapText = true;
                descriptionRange.Interior.Color = CommonHelper.GetColor(CommonHelper.CheckAttributes("BgColor", styleOfDescription, style));
                descriptionRange.Font.Color = CommonHelper.GetColor(CommonHelper.CheckAttributes("FontColor", styleOfDescription, style));
                descriptionRange.Font.Size = Convert.ToDouble(CommonHelper.CheckAttributes("FontSize", styleOfDescription, style));
                descriptionRange.Font.Name = CommonHelper.CheckAttributes("FontFamily", styleOfDescription, style);
                descriptionRange.Borders.Color = XlRgbColor.rgbSlateGray;

                //Get the range of sheet to fill count

                XmlNode styleOfDescTitle = style.SelectSingleNode("ComponentStyle").SelectSingleNode("DescriptionTitle");

                Excel.Range descriptionTitleRange = oSheet1.get_Range("A1", "A1");
                descriptionTitleRange.Value = descriptionTitle.InnerText.Trim();
                //(styleOfDescription.Attributes["FontFamily"] == null || (styleOfDescription.Attributes["FontFamily"].InnerText == null) ? fontFamily : styleOfDescription.Attributes["FontFamily"].InnerText;
                descriptionTitleRange.ColumnWidth = Convert.ToDouble(CommonHelper.CheckAttributes("ColumnWidth", styleOfDescTitle, style));
                descriptionTitleRange.Interior.Color = CommonHelper.GetColor(CommonHelper.CheckAttributes("BgColor", styleOfDescTitle, style));
                descriptionTitleRange.Font.Color = CommonHelper.GetColor(CommonHelper.CheckAttributes("FontColor", styleOfDescTitle, style));
                descriptionTitleRange.Borders.Color = XlRgbColor.rgbSlateGray;
                descriptionTitleRange.Font.Size = Convert.ToDouble(CommonHelper.CheckAttributes("FontSize", styleOfDescTitle, style));
                descriptionTitleRange.Font.Name = CommonHelper.CheckAttributes("FontFamily", styleOfDescTitle, style);
                descriptionTitleRange.Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
                descriptionTitleRange.Font.Bold = true;

                //descriptionTitleRange.Columns.AutoFit();

                oDiscoveryViewook.Save();
                oDiscoveryViewook.Close();

                object misValue = System.Reflection.Missing.Value;
                oBook.Close(false, misValue, misValue);

                oApp.Quit();
                oApp.Application.Quit();

                Marshal.ReleaseComObject(oBook);
                Marshal.ReleaseComObject(workbooks);
                Marshal.ReleaseComObject(oDiscoveryViewook);

                Logger.LogInfoMessage(string.Format("[GeneratePivotReports][GeneratePivotAndSlicersView] Processing Completed for (" + inputCSVFile + ") and sheet " + newSheetName + " is created in Pivot Output file: " + PivotOutputReportFullPath), true);
            }
            catch (Exception ex)
            {
                if (oDiscoveryViewook != null)
                {
                    oDiscoveryViewook.Save();
                    oDiscoveryViewook.Close();
                }

                object misValue = System.Reflection.Missing.Value;
                if (oBook != null)
                {
                    oBook.Close(false, misValue, misValue);
                }
                if (oApp != null)
                {
                    oApp.Quit();
                    oApp.Application.Quit();
                }

                Marshal.ReleaseComObject(oBook);
                Marshal.ReleaseComObject(workbooks);
                Marshal.ReleaseComObject(oDiscoveryViewook);

                Logger.LogErrorMessage(string.Format("[GeneratePivotReports][GeneratePivotAndSlicersView][Exception]: " + ex.Message + ", " + exceptionComment), true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Pivot", ex.Message, ex.ToString(),
                            "[GeneratePivotReports]: GeneratePivotAndSlicersView", ex.GetType().ToString(), exceptionComment);
            }
            finally
            {
                Marshal.ReleaseComObject(oApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
