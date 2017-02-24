using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Collections;
using Microsoft.Office.Interop.Excel;
using System.Xml;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Collections.ObjectModel;
using System.Text.RegularExpressions;

using Microsoft.SharePoint.Client;
using JDP.Remediation.Console.Common.CSV;
using JDP.Remediation.Console.Common.Base;
using JDP.Remediation.Console.Common.Utilities;
using JDP.Remediation.Console.PivotHelper;


namespace JDP.Remediation.Console
{
    public static class GeneratePivotReport
    {
        private static readonly Regex RexCsvSplitter = new Regex(@",(?=(?:[^""]*""[^""]*"")*(?![^""]*""))");

        public static string timeStamp = string.Empty;
        public static string OutputFolderPath = string.Empty;
        public static string outputPath = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
        public static bool flag = false;
        //public static bool flagXML = false;
        public static void DoWork()
        {
            string PivotConfigXMLFileName = string.Empty;
            bool Discovery = false;
            bool PreMigration = false;

            if (!ReadInputOptions(ref Discovery, ref PreMigration))
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                System.Console.WriteLine("Invalid option selected or Exit option is selected. Operation aborted!");
                System.Console.ResetColor();
                return;
            }

            if (Discovery)
            {
                timeStamp = DateTime.Now.ToString("yyyyMMdd_hhmmss");
                try
                {
                    Environment.CurrentDirectory = outputPath;
                    Logger.OpenLog("DT_GeneratePivotReport", timeStamp);
                    PivotConfigXMLFileName = "Discovery-Pivot.xml";

                    //Reading Usage Files Path
                    ReadInputFilesPath(ref OutputFolderPath, PivotConfigXMLFileName, Constants.DTFileName);
                    if (flag == true)
                        return;

                    GeneratePivotReports(OutputFolderPath, PivotConfigXMLFileName, "Component", outputPath);
                }
                catch (Exception ex)
                {
                    Logger.LogErrorMessage(String.Format("[PivotReports: DoWork] failed: Error={0}", ex.Message), true);
                    ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "PivotReports", ex.Message, ex.ToString(), "[PivotReports]: DoWork()", ex.GetType().ToString(), Constants.NotApplicable);
                }
                Logger.CloseLog();
            }

            else if (PreMigration)
            {
                timeStamp = DateTime.Now.ToString("yyyyMMdd_hhmmss");
                try
                {
                    Environment.CurrentDirectory = outputPath;
                    Logger.OpenLog("PreMT_GeneratePivotReport", timeStamp);
                    PivotConfigXMLFileName = "PreMT-Pivot.xml";

                    //Reading Usage Files Path
                    ReadInputFilesPath(ref OutputFolderPath, PivotConfigXMLFileName, Constants.PreMTFileName);
                    if (flag == true)
                        return;

                    GeneratePivotReports(OutputFolderPath, PivotConfigXMLFileName, "Component", outputPath);
                }
                catch (Exception ex)
                {
                    Logger.LogErrorMessage(String.Format("[PivotReports: DoWork] failed: Error={0}", ex.Message), true);
                    ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "PivotReports", ex.Message, ex.ToString(), "[PivotReports]: DoWork()", ex.GetType().ToString(), Constants.NotApplicable);
                }
                Logger.CloseLog();
            }
        }

        private static bool ReadInputOptions(ref bool Discovery, ref bool PreMigration)
        {
            string processOption = string.Empty;
            System.Console.ForegroundColor = System.ConsoleColor.Magenta;
            System.Console.WriteLine("Pivot Reports has to be generated for which of the following files:");
            System.Console.WriteLine("1. Discovery Usage Files");
            System.Console.WriteLine("2. Pre Migration Scan Files");
            System.Console.WriteLine("3. Exit");
            System.Console.ResetColor();
            processOption = System.Console.ReadLine();

            if (processOption.Equals("1"))
                Discovery = true;
            else if (processOption.Equals("2"))
                PreMigration = true;
            else if (processOption.Equals("3"))
                return false;
            else
                return false;

            return true;
        }

        public static void ReadInputFilesPath(ref string OutputFolderPath, string PivotConfigXMLFileName, string fileName)
        {
            System.Console.ForegroundColor = System.ConsoleColor.Cyan;
            Logger.LogMessage(string.Format(@"Enter Complete Path of {0} Files for which Pivot Report has to be generated: (Eg: C:\Test\Files)", fileName), true);
            System.Console.ForegroundColor = System.ConsoleColor.Yellow;
            Logger.LogMessage(string.Format("Note: Before Execution, {0} file has to be saved along with the {1} Files in the same folder", PivotConfigXMLFileName, fileName), true);
            System.Console.ResetColor();
            OutputFolderPath = System.Console.ReadLine();

            //Validating Input File/Folder Path
            if (string.IsNullOrEmpty(OutputFolderPath) || !System.IO.Directory.Exists(OutputFolderPath))
            {
                System.Console.ForegroundColor = System.ConsoleColor.Red;
                Logger.LogErrorMessage("Input files directory is not valid. So, Operation aborted!");
                Logger.LogErrorMessage(@"Please enter path like: E.g. C:\Test\Files");
                System.Console.ResetColor();
                flag = true;
            }
            else
                flag = false;
        }

        public static void GeneratePivotReports(string OutputFolderPath, string PivotConfigXMLFileName, string masterXMLrootNode, string outputPath)
        {
            try
            {
                //Formatting of Output Directory
                FileUtility.ValidateDirectory(ref OutputFolderPath);

                Logger.LogInfoMessage(string.Format("Pivot Report Utility Execution Started for XML: " + PivotConfigXMLFileName), true);
                
                string PivotConfigXMLFilePath = OutputFolderPath + @"\" + PivotConfigXMLFileName;

                string PivotOutputReportFullPath = String.Empty;

                //XmlNodeList otherNodes = null;
                List<XmlNode> otherNodes = new List<XmlNode>();

                if (System.IO.File.Exists(PivotConfigXMLFilePath))
                {
                    //Load Pivot Config Xml
                    XmlNodeList Components = LoadPivotConfigXML(PivotConfigXMLFilePath, masterXMLrootNode, OutputFolderPath, ref PivotOutputReportFullPath, ref otherNodes, outputPath);

                    if (Components != null)
                    {
                        GeneratePivotReportsUsingCSVFiles(Components, OutputFolderPath, PivotOutputReportFullPath, otherNodes);
                    }
                }
                else
                {
                    string ErrorMessage = "[GeneratePivotReports] File Not Found Error: XML file (" + PivotConfigXMLFileName + ") is not present in path:" + PivotConfigXMLFilePath;
                    Logger.LogErrorMessage(ErrorMessage, true);
                    ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Pivot", ErrorMessage, Constants.NotApplicable,
                            "GeneratePivotReports", Constants.NotApplicable, ErrorMessage);
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(string.Format("[GeneratePivotReports][Exception]: " + ex.Message), true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Pivot", ex.Message, ex.ToString(),
                            "GeneratePivotReports", ex.GetType().ToString(), Constants.NotApplicable);
            }
            finally
            {
                System.Console.ForegroundColor = System.ConsoleColor.Green;
                Logger.LogSuccessMessage(string.Format("Pivot Report Utility Execution Completed Successfully & Logger, Exception and Output files are saved in path: " + outputPath), true);
                System.Console.ResetColor();
            }
        }
        public static XmlNodeList LoadPivotConfigXML(string PivotConfigXMLFilePath, string masterXMLrootNode, string OutputFolderPath, ref string PivotOutputReportFullPath, ref List<XmlNode> otherNodes, string outputPath)
        {
            XmlNodeList Components = null;

            try
            {
                //[START] Read/Load Pivot XML Config File
                Logger.LogInfoMessage(string.Format("[GeneratePivotReports][LoadPivotConfigXML] Pivot Config XML(" + PivotConfigXMLFilePath + ") loading process has been initiated"), false);
                var xDoc = new XmlDocument();
                xDoc.Load(PivotConfigXMLFilePath);

                var root = xDoc.DocumentElement;

                if (root != null)
                {
                    Components = root.SelectNodes(masterXMLrootNode);

                    if (Components != null)
                    {
                        //Pivot Output Report File Name
                        var PivotReportOutput = root.SelectSingleNode("PivotReportOutput");
                        string PivotOutputReportFileName = PivotReportOutput.InnerText;

                        PivotOutputReportFullPath = outputPath + @"\" + PivotOutputReportFileName;

                        otherNodes.Add(root.SelectSingleNode("SummaryView"));
                        otherNodes.Add(root.SelectSingleNode("TOC"));
                        otherNodes.Add(root.SelectSingleNode("Style"));
                        otherNodes.Add(root.SelectSingleNode("SourceFile"));

                        //Delete OLD/Existing Pivot Report
                        if (System.IO.File.Exists(PivotOutputReportFullPath))
                        {
                            System.IO.File.Delete(PivotOutputReportFullPath);
                        }
                        //Delete OLD/Existing Pivot Report

                        Excel.Application excelApplication = new Excel.Application();
                        Excel.Workbook excelWorkBook = excelApplication.Workbooks.Add();

                        try
                        {
                            //Creating TOC Sheet - Table Of Content
                            Excel.Worksheet excelWorkSheet = (Excel.Worksheet)excelWorkBook.ActiveSheet;
                            excelWorkSheet.Name = otherNodes.Find(item => item.Name == "TOC").Name;
                            excelApplication.SheetsInNewWorkbook = 1;
                            excelWorkBook.SaveAs(PivotOutputReportFullPath, Excel.XlFileFormat.xlOpenXMLWorkbook, Missing.Value, Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlUserResolution, true, Missing.Value, Missing.Value, Missing.Value);
                            excelWorkBook.Close();

                            excelApplication.Application.Quit();
                            excelApplication.Quit();

                            Marshal.ReleaseComObject(excelWorkSheet);
                            Marshal.ReleaseComObject(excelWorkBook);
                            Marshal.ReleaseComObject(excelApplication);

                            excelApplication = null;
                        }
                        catch (Exception ex)
                        {
                            if (excelWorkBook != null)
                            {
                                excelWorkBook.Close();
                            }

                            if (excelApplication != null)
                            {
                                excelApplication.Quit();
                                excelApplication.Application.Quit();
                            }

                            Marshal.ReleaseComObject(excelWorkBook);

                            Logger.LogErrorMessage(string.Format("[GeneratePivotReports][LoadPivotConfigXML][Exception]: " + ex.Message + "\n" + ex.StackTrace.ToString()), true);
                            ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Pivot", ex.Message, ex.ToString(),
                            "[GeneratePivotReports]: LoadPivotConfigXML", ex.GetType().ToString(), "PivotConfigXMLFilePath: " + PivotConfigXMLFilePath);
                        }
                        finally
                        {
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            Marshal.ReleaseComObject(excelWorkBook);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(string.Format("[GeneratePivotReports][LoadPivotConfigXML][Exception]: " + ex.Message), true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Pivot", ex.Message, ex.ToString(),
                            "[GeneratePivotReports]: GeneratePivotAndSlicersView", ex.GetType().ToString(), "PivotConfigXMLFilePath: " + PivotConfigXMLFilePath);
            }

            Logger.LogInfoMessage(string.Format("[GeneratePivotReports][LoadPivotConfigXML] Pivot Config XML(" + PivotConfigXMLFilePath + ") loading process has been completed"), true);

            return Components;
        }

        public static void GeneratePivotReportsUsingCSVFiles(XmlNodeList Components, string InputCSVFolderPath, string PivotOutputReportFullPath, List<XmlNode> otherNodes)
        {
            try
            {
                //Summavry View Components
                Dictionary<string, string[]> dSummaryViewComponents = new Dictionary<string, string[]>();

                //Hyperlinks for Components and Summary
                List<Hashtable> htHyperLinks = new List<Hashtable>();
                Hashtable htHyperLinksForComponents = new Hashtable();
                Hashtable htHyperLinksForSummary = new Hashtable();

                //Populating the SummaryView HyperLink hashtable with Name and Description
                XmlNode summaryViewNode = otherNodes.Find(item => item.Name == "SummaryView");
                htHyperLinksForSummary.Add(summaryViewNode.Name, summaryViewNode.SelectSingleNode("Description").InnerText);

                //Generate Pivot Views for All Components by reading the Component Tag from Pivot Config Xml
                //Components.

                foreach (XmlNode component in Components)
                {
                    GeneratePivotReportForMultipleFiles(InputCSVFolderPath, component, PivotOutputReportFullPath, ref dSummaryViewComponents, ref htHyperLinksForComponents, otherNodes);
                }

                SummaryViewHelper.ComponentsSummaryView(summaryViewNode, dSummaryViewComponents, PivotOutputReportFullPath);
                htHyperLinks.Add(htHyperLinksForComponents);
                htHyperLinks.Add(htHyperLinksForSummary);
                //Create Summary View  Report

                //Apply Page Index and Back To Page Key will be Component/Sheet Name and value will position of text to display
                CommonHelper.Create_HyperLinks(PivotOutputReportFullPath, htHyperLinks, otherNodes);
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(string.Format("[GeneratePivotReports][GeneratePivotReportsUsingCSVFiles][Exception]: " + ex.Message), true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Pivot", ex.Message, ex.ToString(),
                            "[GeneratePivotReports]: GeneratePivotReportsUsingCSVFiles", ex.GetType().ToString(), Constants.NotApplicable);
            }
        }

        /// <summary>
        /// Iterating through all the components for slicers, pivot, row fields, and value columns and collecting them in Hash Tables
        /// Passing them to other method wherein all the pivot tables, slicers, and row fields, filters are created based on values from 
        /// their hashTables.
        /// </summary>
        /// <param name="InputCSVFolderPath"></param>
        /// <param name="component"></param>
        /// <param name="PivotOutputReportFullPath"></param>
        /// <param name="dSummaryViewComponents"></param>
        /// <param name="htHyperLinks"></param>
        /// <param name="otherNodes"></param>
        private static void GeneratePivotReportForMultipleFiles(
        string InputCSVFolderPath,
        XmlNode component, string PivotOutputReportFullPath,
        ref Dictionary<string, string[]> dSummaryViewComponents,
        ref Hashtable htHyperLinks,
        List<XmlNode> otherNodes)
        {
            StringBuilder exceptionCommentsInfo = new StringBuilder();
            int numberOfFilesCount = 1;
            string sheetName = string.Empty;
            try
            {
                //Pivot Report - RowFields
                Hashtable htRowPivotfields;
                //Pivot Report - PageFilters
                Hashtable htPivotPageFilters;
                //Pivot Report - PageSlicers
                Hashtable htPageSlicers;

                //Slicers and Dicers HashTables
                Hashtable htComponentSliceandDiceViews = new Hashtable();

                int RowPivotfieldCount = 1, PivotPageFiltersCount = 1;
                int sliceandDiceViewsCount = 1;

                //Get the Components Lists by Reading the Component Tags and Attribute. 
                //Example Content Types, Master Pages, etc from Pivot Config XML
                string componentName = component.Attributes["Name"].InnerText;
                string InputFileName = component.Attributes["InputFileName"].InnerText;
                string summaryViewColumn = component.Attributes["SummaryViewColumn"].InnerText;
                string description = component.SelectSingleNode("Description").InnerText;
                string componentSliceAndDiceSheetname = "";

                //Exception Comments
                exceptionCommentsInfo.Append("ComponentName: " + componentName + ", InputFileName: " + InputFileName + ", SummaryViewColumn: " + summaryViewColumn);

                //Read All The INput File CSV for a Component. If Any Components has the multiple Usage or Input File, we are reading all files
                //Example: If Content Type Usage Files are - ContentType_Usage.csv, ContentType_Usage_03112016_035049.csv
                string searchpattern = System.IO.Path.GetFileNameWithoutExtension(InputCSVFolderPath + "\\" + InputFileName);
                searchpattern = searchpattern + "*.csv";
                string[] files = FileUtility.FindAllFilewithSearchPattern(InputCSVFolderPath, searchpattern);

                //Null Check
                if (files != null)
                {
                    if (files.Count() > 0)
                    {
                        foreach (string filePath in files)
                        {
                            string inputCSVFile = filePath;
                            string inputFileName = filePath.Substring(filePath.LastIndexOf("\\") + 1);

                            htRowPivotfields = new Hashtable();
                            htPivotPageFilters = new Hashtable();
                            htPageSlicers = new Hashtable();

                            if (System.IO.File.Exists(inputCSVFile))
                            {
                                //Get Data InputDataSheetName
                                string inputDataSheetName = System.IO.Path.GetFileNameWithoutExtension(filePath);

                                //Deleting columns from csv
                                if (inputDataSheetName == "CThavingFeatureIDTag_Definition_Usage" || inputDataSheetName == "ListTemplates_Usage" || inputDataSheetName == "PreMT_CThavingFeatureIDTag_Definition")
                                {
                                    deleteColumnFromCSV(inputDataSheetName, InputCSVFolderPath, inputCSVFile, inputFileName);
                                }

                                exceptionCommentsInfo.Clear();
                                exceptionCommentsInfo.Append(", InputDataSheetName: " + inputDataSheetName);

                                Excel.Application oApp;
                                Excel.Worksheet oSheet;
                                Excel.Workbook oBook = null;

                                oApp = new Excel.Application();

                                try
                                {
                                    oBook = oApp.Workbooks.Open(inputCSVFile);
                                    if (inputDataSheetName.Length >= Constants.SheetNameMaxLength)
                                        oSheet = (Excel.Worksheet)oBook.Sheets.get_Item(1);
                                    else
                                        oSheet = (Excel.Worksheet)oBook.Sheets.get_Item(inputDataSheetName);

                                    // Now capture range of the first sheet 
                                    Excel.Range oRange = oSheet.UsedRange;

                                    if (oRange.Rows.Count > 1)
                                    {
                                        if (files.Count() > 1)
                                        {
                                            sheetName = componentName + "_" + numberOfFilesCount;
                                            //SliceAndDiceSheet
                                            componentSliceAndDiceSheetname = componentName + "_" + numberOfFilesCount + Constants.SliceAndDiceSheet_Suffix;
                                        }
                                        else
                                        {
                                            sheetName = componentName;
                                            componentSliceAndDiceSheetname = componentName + Constants.SliceAndDiceSheet_Suffix;
                                        }

                                        dSummaryViewComponents.Add(sheetName, new string[] { inputCSVFile, summaryViewColumn });

                                        string pivotCountField = string.Empty; string pivotSliceandDiceCountField = string.Empty;
                                        string slicerStyle = string.Empty;

                                        //Read Item Tags inside the Component Tag, to generate - Pivot and Slice and Dice View
                                        foreach (XmlNode ItemNode in component.SelectNodes("Item"))
                                        {
                                            string itemType = ItemNode.Attributes["Type"].InnerText;

                                            //Create Pivot View
                                            if (itemType == "PivotView")
                                            {
                                                var count = ItemNode.SelectSingleNode("ValueColumn");
                                                pivotCountField = count.Attributes["Name"].InnerText;

                                                var slicerstyle = ItemNode.SelectSingleNode("SlicersStyling");
                                                slicerStyle = slicerstyle.Attributes["Style"].InnerText;

                                                //Row Filter
                                                foreach (XmlNode RowFieldRoot in ItemNode.SelectNodes("Rows"))
                                                {
                                                    foreach (XmlNode rowFeild in RowFieldRoot.ChildNodes)
                                                    {
                                                        string rowPageFieldName = rowFeild.Attributes["Column"].InnerText;
                                                        string rowPageLabel = rowFeild.Attributes["Label"].InnerText;
                                                        htRowPivotfields.Add(RowPivotfieldCount, rowPageFieldName + "~" + rowPageLabel);
                                                        RowPivotfieldCount++;
                                                    }
                                                }

                                                //Page Filters
                                                foreach (XmlNode FilterFeildRoot in ItemNode.SelectNodes("Filters"))
                                                {
                                                    foreach (XmlNode filterFeild in FilterFeildRoot.ChildNodes)
                                                    {
                                                        string filterName = filterFeild.InnerText;
                                                        htPivotPageFilters.Add(PivotPageFiltersCount, filterName);
                                                        PivotPageFiltersCount++;
                                                    }
                                                }

                                                //Slicers
                                                foreach (XmlNode SlicersRoot in ItemNode.SelectNodes("Slicers"))
                                                {
                                                    int i = 0;
                                                    i = SlicersRoot.ChildNodes.Count - 1;
                                                    foreach (XmlNode SlicersRootFeild in SlicersRoot.ChildNodes)
                                                    {
                                                        string slicerName = SlicersRootFeild.InnerText;
                                                        htPageSlicers.Add(i, slicerName + "~" + slicerStyle);
                                                        i--;
                                                    }
                                                }
                                            }

                                            //Create Slice and Dice View Sheet
                                            if (itemType == "SliceDiceView")
                                            {
                                                //Loop views for each Component
                                                foreach (XmlNode viewsNode in ItemNode.SelectNodes("Views"))
                                                {
                                                    foreach (XmlNode viewNode in viewsNode.ChildNodes)
                                                    {
                                                        htComponentSliceandDiceViews.Add(inputDataSheetName + sliceandDiceViewsCount, viewNode.OuterXml);
                                                        sliceandDiceViewsCount++;
                                                    }
                                                }
                                            }
                                        }

                                        //Pivot Sheet
                                        string componentPivotSheetName = sheetName;
                                        //Length of Sheet Name Should be less than 31 Char
                                        if (componentPivotSheetName.Length >= Constants.SheetNameMaxLength)
                                        {
                                            componentPivotSheetName = componentPivotSheetName.Substring(0, Constants.SheetNameMaxLength);
                                        }


                                        //Length of Sheet Name Should be less than 31 Char
                                        if (componentSliceAndDiceSheetname.Length >= Constants.SheetNameMaxLength)
                                        {
                                            componentSliceAndDiceSheetname = componentSliceAndDiceSheetname.Substring(0, Constants.SheetNameMaxLength);
                                        }

                                        //Create Pivot View and Slicer View Sheet
                                        PivotViewHelper.GeneratePivotAndSlicersView(inputCSVFile, PivotOutputReportFullPath, ref componentPivotSheetName, inputDataSheetName, componentPivotSheetName, htRowPivotfields, htPivotPageFilters, htPageSlicers, pivotCountField, component, numberOfFilesCount, otherNodes);

                                        htHyperLinks.Add(sheetName + "~" + inputFileName, description);

                                        RowPivotfieldCount = 1;
                                        PivotPageFiltersCount = 1;
                                        htPageSlicers = null;
                                        htPivotPageFilters = null;
                                        htRowPivotfields = null;
                                        sliceandDiceViewsCount = 1;
                                        numberOfFilesCount += 1;
                                    }
                                    else
                                    {
                                        Logger.LogInfoMessage(string.Format("[GeneratePivotReports][GeneratePivotReportForMultipleFiles]" + " No records available for the component " + componentName + " in file " + inputDataSheetName), true);
                                    }

                                    object misValue = System.Reflection.Missing.Value;
                                    oBook.Close(false, misValue, misValue);

                                    oApp.Application.Quit();
                                    oApp.Quit();

                                    Marshal.ReleaseComObject(oSheet);
                                    Marshal.ReleaseComObject(oBook);

                                }
                                catch (Exception ex)
                                {
                                    if (oBook != null)
                                    {
                                        oBook.Close();
                                    }

                                    if (oApp != null)
                                    {
                                        oApp.Quit();
                                        oApp.Application.Quit();
                                    }

                                    Marshal.ReleaseComObject(oBook);

                                    Logger.LogErrorMessage(string.Format("[GeneratePivotReports][GeneratePivotReportForMultipleFiles][Exception]: " + ex.Message), true);
                                    ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Pivot", ex.Message, ex.ToString(),
                                                "[GeneratePivotReports]: GeneratePivotReportForMultipleFiles", ex.GetType().ToString(), "ExceptionCommentsInfo: " + exceptionCommentsInfo);
                                }
                                finally
                                {
                                    Marshal.ReleaseComObject(oApp);
                                    oApp = null;

                                    GC.Collect();
                                    GC.WaitForPendingFinalizers();
                                }

                                if (inputDataSheetName == "CThavingFeatureIDTag_Definition_Usage" || inputDataSheetName == "ListTemplates_Usage" || inputDataSheetName == "PreMT_CThavingFeatureIDTag_Definition")
                                {

                                    string path1 = InputCSVFolderPath + @"\" + "Backup";
                                    //copy file to backup folder
                                    string destCSVFile1 = path1 + @"\" + inputFileName;
                                    System.IO.File.Copy(destCSVFile1, inputCSVFile, true);

                                    DeleteFolderAndFiles(path1);
                                }
                            }
                            else
                            {
                                string ErrorMessage = "[GeneratePivotReports][GeneratePivotReportForMultipleFiles] File Not Found Error: Input CSV file is not present in path:" + inputCSVFile + ", ExceptionCommentsInfo: " + exceptionCommentsInfo;

                                Logger.LogErrorMessage(ErrorMessage, true);
                                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Pivot", "File Not Found", ErrorMessage,
                                            "[GeneratePivotReports]: GeneratePivotReportForMultipleFiles", "File Not Found", "Input CSV File (" + inputCSVFile + ")");
                            }
                        }
                    }
                    else
                    {
                        string ErrorMessage = "[GeneratePivotReports][GeneratePivotReportForMultipleFiles] File Not Found Error: Input CSV file is not present in path:" + InputCSVFolderPath + "\\" + InputFileName + ", " + exceptionCommentsInfo;
                        Logger.LogErrorMessage(ErrorMessage, true);
                        ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Pivot", "File Not Found", ErrorMessage,
                                            "[GeneratePivotReports]: GeneratePivotReportForMultipleFiles", "File Not Found", "Input CSV File (" + InputCSVFolderPath + "\\" + InputFileName + ")");
                    }
                }
                else
                {
                    string ErrorMessage = "[GeneratePivotReports][GeneratePivotReportForMultipleFiles] File Not Found Error: Input CSV file is not present in path InputCSVFolderPath: " + InputCSVFolderPath + ", SearchPattern: " + searchpattern + ", " + exceptionCommentsInfo;
                    Logger.LogErrorMessage(ErrorMessage, true);
                }

            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(string.Format("[GeneratePivotReports][GeneratePivotReportForMultipleFiles][Exception]: " + ex.Message + ", ExceptionCommentsInfo: " + exceptionCommentsInfo), true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Pivot", ex.Message, ex.ToString(),
                                            "[GeneratePivotReports]: GeneratePivotReportForMultipleFiles", ex.GetType().ToString(), exceptionCommentsInfo.ToString());
            }
        }

        public static System.Data.DataTable csvToDataTable(string file, bool isRowOneHeader)
        {

            System.Data.DataTable csvDataTable = new System.Data.DataTable();

            //no try/catch - add these in yourselfs or let exception happen
            String[] csvData = System.IO.File.ReadAllLines(file);

            //if no data in file ‘manually’ throw an exception
            if (csvData.Length == 0)
            {
                throw new Exception("CSV File Appears to be Empty");
            }

            String[] headings = csvData[0].Split(',');
            int index = 0; //will be zero or one depending on isRowOneHeader

            if (isRowOneHeader) //if first record lists headers
            {
                index = 1; //so we won’t take headings as data

                //for each heading
                for (int i = 0; i < headings.Length; i++)
                {
                    //replace spaces with underscores for column names
                    headings[i] = headings[i].Replace(" ", "_");
                    //add a column for each heading
                    csvDataTable.Columns.Add(headings[i], typeof(string));
                }
            }
            else //if no headers just go for col1, col2 etc.
            {
                for (int i = 0; i < headings.Length; i++)
                {
                    //create arbitary column names
                    csvDataTable.Columns.Add("col" + (i + 1).ToString(), typeof(string));
                }
            }

            //populate the DataTable
            for (int i = index; i < csvData.Length; i++)
            {
                //create new rows
                System.Data.DataRow row = csvDataTable.NewRow();
                for (int j = 0; j < headings.Length; j++)
                {
                    //fill them
                    //row[j] = csvData[i].Split(',')[j];
                    row[j] = RexCsvSplitter.Split(csvData[i])[j];
                }
                //add rows to over DataTable
                csvDataTable.Rows.Add(row);
            }
            //return the CSV DataTable
            return csvDataTable;
        }

        public static void WriteToCsvFile(this System.Data.DataTable dataTable, string filePath)
        {
            StringBuilder fileContent = new StringBuilder();

            foreach (var col in dataTable.Columns)
            {
                fileContent.Append(col.ToString() + ",");
            }
            fileContent.Replace(",", System.Environment.NewLine, fileContent.Length - 1, 1);

            foreach (System.Data.DataRow dr in dataTable.Rows)
            {
                foreach (var column in dr.ItemArray)
                {
                    fileContent.Append("\"" + column.ToString() + "\",");
                }
                fileContent.Replace(",", System.Environment.NewLine, fileContent.Length - 1, 1);
            }
            System.IO.File.WriteAllText(filePath, fileContent.ToString());
        }

        public static void deleteColumnFromCSV(string sheetName, string inputCSVFolderPath, string inputCSVFile, string inputFileName)
        {
            try
            {
                //creating backup folder
                string path = inputCSVFolderPath + @"\" + "Backup";
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                //copy file to backup folder
                string destCSVFile = path + @"\" + inputFileName;
                System.IO.File.Copy(inputCSVFile, destCSVFile, true);

                System.Data.DataTable newDataTable = new System.Data.DataTable();
                newDataTable = csvToDataTable(inputCSVFile, true);

                //Removing Unwanted column
                if (sheetName == "CThavingFeatureIDTag_Definition_Usage" || sheetName == "PreMT_CThavingFeatureIDTag_Definition")
                {
                    newDataTable.Columns.Remove("Definition");
                }
                else if (sheetName == "ListTemplates_Usage")
                {
                    newDataTable.Columns.Remove("AssociatedContentTypes");
                }

                WriteToCsvFile(newDataTable, inputCSVFile);
            }
            catch (Exception ex)
            {
                string ErrorMessage = "[GeneratePivotReports][GeneratePivotReportForMultipleFiles][DeleteColumnFromCSV] Error While deleting columns from CSV File";
                Logger.LogErrorMessage(ErrorMessage, true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Pivot", ex.Message, ErrorMessage,
                           "[GeneratePivotReports]: DeleteColumnFromCSV", ex.GetType().ToString(), "Error While deleting columns from CSV File");
            }
        }

        public static void DeleteFolderAndFiles(string pathToExtractWSP)
        {
            if (Directory.Exists(pathToExtractWSP))
            {
                Directory.SetCurrentDirectory(pathToExtractWSP);
                DirectoryInfo directory = new DirectoryInfo(pathToExtractWSP);
                try
                {
                    foreach (System.IO.FileInfo file in directory.GetFiles())
                    {
                        try
                        {
                            file.Delete();
                        }
                        catch { }
                    }
                    foreach (System.IO.DirectoryInfo subDirectory in directory.GetDirectories())
                    {
                        try
                        {
                            subDirectory.Delete(true);
                        }
                        catch { }
                    }
                }
                catch
                { }
            }
        }
    }
}
