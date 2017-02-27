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
using Microsoft.SharePoint.Client;
using System.Xml;
using System.Runtime.InteropServices;
using Marshal = System.Runtime.InteropServices.Marshal;
using System.Management;
using System.Diagnostics;
using System.Drawing;
using System.Security;


using JDP.Remediation.Console.Common.CSV;
using JDP.Remediation.Console.Common.Base;
using JDP.Remediation.Console.Common.Utilities;

namespace JDP.Remediation.Console.PivotHelper
{
    class CommonHelper
    {
        
        /// <summary>
        /// Creating Hyper Links of every component in TOC sheet and adding corresponding description against them.
        /// Also Adding Back To Index links and Source file Name texts in every sheets.
        /// </summary>
        /// <param name="PivotOutputReportFullPath"></param>
        /// <param name="htHyperLinks"></param>
        /// <param name="otherNodes"></param>
        public static void Create_HyperLinks(string PivotOutputReportFullPath, List<Hashtable> htHyperLinks, List<XmlNode> otherNodes)
        {
            // Creates a new Excel Application
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = null;
            var workbooks = excelApp.Workbooks;

            XmlNode tocNode = otherNodes.Find(item => item.Name == "TOC");
            XmlNode sourceNode = otherNodes.Find(item => item.Name == "SourceFile");
            XmlNode style = otherNodes.Find(item => item.Name == "Style");

            Logger.LogInfoMessage(string.Format("[GeneratePivotReports][Create_HyperLinks] Processing Started to add Links in Table of Content Sheet and Back Links in all output sheets"), false);

            try
            {
                excelWorkbook = workbooks.Open(PivotOutputReportFullPath);
            }
            catch
            {
                //Create a new workbook if the existing workbook failed to open.
                excelWorkbook = excelApp.Workbooks.Add();
            }

            try
            {
                // The following gets the Worksheets collection
                Excel.Sheets excelSheets = excelWorkbook.Worksheets;
                XmlNode tocHeading = tocNode.SelectSingleNode("TOCHeading");
                XmlNode tocTitle = tocNode.SelectSingleNode("TOCTitle");
                XmlNode tocDescription = tocNode.SelectSingleNode("TOCDescription");
                XmlNode sourceHead = sourceNode.SelectSingleNode("SourceFileHeading");
                XmlNode backToIndex = sourceNode.SelectSingleNode("BackToIndex");

                XmlNode tocHeadStyle = style.SelectSingleNode("TOCStyle").SelectSingleNode("TOCHeading");
                XmlNode tocTitleStyle = style.SelectSingleNode("TOCStyle").SelectSingleNode("TOCTitle");
                XmlNode tocDescStyle = style.SelectSingleNode("TOCStyle").SelectSingleNode("TOCDescription");
                XmlNode tocStyle = style.SelectSingleNode("TOCStyle").SelectSingleNode("Style");
                XmlNode sourceHeadStyle = style.SelectSingleNode("SourceFileStyle").SelectSingleNode("SourceFileHeading");
                XmlNode sourceFileNameStyle = style.SelectSingleNode("SourceFileStyle").SelectSingleNode("SourceFileName");
                XmlNode backToIndexStyle = style.SelectSingleNode("SourceFileStyle").SelectSingleNode("BackToIndex");

                string sheetName = "";
                string sourceFileNameText = "";

                // The following gets Sheet1 for editing
                string currentSheet = tocNode.Name;
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);

                //Fixed Table Location
                Excel.Range tRange = excelWorksheet.get_Range("A2", "A22");
                //borders.Weight = 1d;
                int rowNumber = 3;

                foreach (var htRowIndex in htHyperLinks)
                {
                    foreach (DictionaryEntry hyperlink in htRowIndex.Cast<DictionaryEntry>().OrderBy(item => item.Key).ToList())
                    {
                        if (!htRowIndex.Keys.Cast<String>().Contains("SummaryView"))
                        {
                            string[] keyValue = hyperlink.Key.ToString().Split('~');
                            sheetName = keyValue[0];
                            sourceFileNameText = keyValue[1];
                            //to put the header for the table of contents sheet
                            if (rowNumber == 3)
                            {
                                Excel.Range excelCellRowHeader = (Excel.Range)excelWorksheet.get_Range("A2", "B2");
                                excelCellRowHeader.Merge(Missing.Value);
                                excelCellRowHeader.Value = tocHeading.InnerText;
                                excelCellRowHeader.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                excelCellRowHeader.Font.Size = Convert.ToDouble(CheckAttributes("FontSize", tocHeadStyle, style));
                                excelCellRowHeader.Font.Color = CommonHelper.GetColor(CheckAttributes("FontColor", tocHeadStyle, style));
                                Excel.Borders border = excelCellRowHeader.Borders;
                                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                                excelCellRowHeader.Columns.AutoFit();
                                excelCellRowHeader.Font.Bold = true;
                                excelCellRowHeader.Font.Underline = true;
                                excelCellRowHeader.Font.Name = CheckAttributes("FontFamily", tocHeadStyle, style);
                                rowNumber++;
                            }

                            if (rowNumber == 4)
                            {
                                var excelCellRowHeader2 = (Excel.Range)excelWorksheet.get_Range("A3");
                                excelCellRowHeader2.Value = tocTitle.InnerText;
                                excelCellRowHeader2.Font.Color = CommonHelper.GetColor(CheckAttributes("FontColor", tocTitleStyle, style));
                                Excel.Borders border2 = excelCellRowHeader2.Borders;
                                border2.LineStyle = Excel.XlLineStyle.xlContinuous;
                                excelCellRowHeader2.Columns.AutoFit();
                                excelCellRowHeader2.Font.Bold = true;
                                excelCellRowHeader2.Font.Size = Convert.ToDouble(CheckAttributes("FontSize", tocTitleStyle, style));
                                excelCellRowHeader2.Font.Name = CheckAttributes("FontFamily", tocTitleStyle, style);
                                excelCellRowHeader2.Interior.Color = CommonHelper.GetColor(CheckAttributes("BgColor", tocTitleStyle, style));

                                Excel.Range excelCellRowHeader1 = (Excel.Range)excelWorksheet.get_Range("B3");
                                excelCellRowHeader1.Value = tocDescription.InnerText;

                                excelCellRowHeader1.ColumnWidth = Convert.ToDouble(CheckAttributes("ColumnWidth", tocDescStyle, style));
                                excelCellRowHeader1.Font.Color = CommonHelper.GetColor(CheckAttributes("FontColor", tocDescStyle, style));
                                Excel.Borders border1 = excelCellRowHeader1.Borders;
                                border1.LineStyle = Excel.XlLineStyle.xlContinuous;
                                excelCellRowHeader1.Font.Bold = true;
                                excelCellRowHeader1.Interior.Color = CommonHelper.GetColor(CheckAttributes("BgColor", tocDescStyle, style));
                                excelCellRowHeader1.Font.Size = Convert.ToDouble(CheckAttributes("FontSize", tocDescStyle, style));
                                excelCellRowHeader1.Font.Name = CheckAttributes("FontFamily", tocDescStyle, style);
                                rowNumber++;
                            }
                        }
                        else
                        {
                            sheetName = hyperlink.Key.ToString();
                            rowNumber = 4;
                        }

                        // The following gets cell A1 for editing
                        Excel.Range excelCell = (Excel.Range)excelWorksheet.get_Range("A" + rowNumber);
                        excelWorksheet.Activate();
                        //Add the Text for hyper Link in Table of contents Sheet     
                        excelCell.Value = sheetName;
                        //var s = tocStyle.Attributes["BgColor"];
                        //excelCell.Font.Color = CommonHelper.GetColor((tocStyle.Attributes["BgColor"].InnerText == null) ? bgColor : tocStyle.Attributes["BgColor"].InnerText);
                        Excel.Borders borders = excelCell.Borders;
                        borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        excelCell.Font.Bold = true;
                        excelWorksheet.Hyperlinks.Add(excelCell, "#" + sheetName + "!A1", Type.Missing, Type.Missing, sheetName);

                        excelWorksheet.Application.Range["A" + rowNumber].Select();
                        //excelApp.Selection.Font;
                        excelWorksheet.Application.Selection.Font.Name = CheckAttributes("FontFamily", tocStyle, style);
                        excelWorksheet.Application.Selection.Font.Size = Convert.ToDouble(CheckAttributes("FontSize", tocStyle, style));

                        Excel.Range excelCell3 = (Excel.Range)excelWorksheet.get_Range("B" + rowNumber);
                        //Add the Text for hyper Link in Table of contents Sheet
                        excelCell3.Value = hyperlink.Value.ToString().Trim();
                        excelCell3.Font.Color = CommonHelper.GetColor(CheckAttributes("FontColor", tocStyle, style));
                        Excel.Borders borders2 = excelCell3.Borders;
                        borders2.LineStyle = Excel.XlLineStyle.xlContinuous;
                        excelCell3.Font.Size = Convert.ToDouble(CheckAttributes("FontSize", tocStyle, style));
                        excelCell3.WrapText = true;

                        Excel.Worksheet excelWorksheet2 = (Excel.Worksheet)excelSheets.get_Item(Convert.ToString(sheetName));
                        excelWorksheet2.Activate();
                        Range Line = (Range)excelWorksheet2.Rows[1];
                        Line.Insert();
                        Excel.Range excelCell2 = (Excel.Range)excelWorksheet2.get_Range(Convert.ToString("A1"));
                        excelCell2.Value = backToIndex.InnerText;
                        excelWorksheet2.Hyperlinks.Add(excelCell2, "#" + currentSheet + "!A1", Type.Missing, Type.Missing, Type.Missing);

                        excelWorksheet2.Application.Range["A1"].Select();
                        //excelApp.Selection.Font;
                        excelWorksheet2.Application.Selection.Font.Name = CheckAttributes("FontFamily", backToIndexStyle, style);
                        excelWorksheet2.Application.Selection.Font.Size = Convert.ToDouble(CheckAttributes("FontSize", backToIndexStyle, style));

                        if (!sheetName.Equals("SummaryView"))
                        {
                            //Get the range of sheet to fill count
                            Excel.Range sourceFileTitle = excelWorksheet2.get_Range("B1", "B1");
                            sourceFileTitle.Value = sourceHead.InnerText.Trim();
                            sourceFileTitle.Font.Color = CommonHelper.GetColor(CheckAttributes("FontColor", sourceHeadStyle, style));
                            sourceFileTitle.Font.Size = Convert.ToDouble(CheckAttributes("FontSize", sourceHeadStyle, style));
                            sourceFileTitle.Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
                            sourceFileTitle.Font.Bold = true;
                            sourceFileTitle.Columns.AutoFit();

                            //Get the range of sheet to fill count
                            Excel.Range sourceFileName = excelWorksheet2.get_Range("C1", "F1");
                            sourceFileName.Merge();
                            sourceFileName.Value = sourceFileNameText;
                            sourceFileName.Font.Color = CommonHelper.GetColor(CheckAttributes("FontColor", sourceFileNameStyle, style));
                            sourceFileName.Font.Size = Convert.ToDouble(CheckAttributes("FontSize", sourceFileNameStyle, style));
                            sourceFileName.Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
                            sourceFileName.Columns.AutoFit();
                        }
                        rowNumber++;
                    }
                }
                tRange.Columns.AutoFit();

                // Close the excel workbook
                excelWorkbook.Close(true, Type.Missing, Type.Missing);
                workbooks.Close();
                excelApp.Application.Quit();
                excelApp.Quit();
                Marshal.ReleaseComObject(excelSheets);
                Marshal.ReleaseComObject(excelWorkbook);
                Marshal.ReleaseComObject(workbooks);

                Logger.LogInfoMessage(string.Format("[GeneratePivotReports][Create_HyperLinks] Process Completed to add Links in Table of Content Sheet and Back Links in all output sheets"), true);
            }
            catch (Exception ex)
            {
                if (excelWorkbook != null)
                {
                    excelWorkbook.Save();
                    excelWorkbook.Close();
                }

                if (excelApp != null)
                {
                    excelApp.Quit();
                    excelApp.Application.Quit();
                }

                Marshal.ReleaseComObject(workbooks);
                Marshal.ReleaseComObject(excelWorkbook);

                Logger.LogErrorMessage(string.Format("[GeneratePivotReports][Create_HyperLinks][Exception]: " + ex.Message + "\n" + ex.StackTrace.ToString()), true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Pivot", ex.Message, ex.ToString(),
                            "[GeneratePivotReports]: Create_HyperLinks", ex.GetType().ToString(), Constants.NotApplicable);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Marshal.ReleaseComObject(excelApp);
            }
        }

        /// <summary>
        /// In this method we are checking the font attributes of the respective nodes. If they are not present, or their value is null
        /// then we are assigning the default attributes of Style tag in their place.
        /// </summary>
        /// <param name="attribute"></param>
        /// <param name="destNode"></param>
        /// <param name="defaultStyleNode"></param>
        /// <returns></returns>
        public static string CheckAttributes(string attribute, XmlNode destNode, XmlNode defaultStyleNode)
        {
            var returnValue = "";
            var defaultValue = defaultStyleNode.Attributes["Default" + attribute].InnerText;
            if (destNode != null)
            {
                returnValue = (destNode.Attributes[attribute] == null) ? defaultValue : (destNode.Attributes[attribute].InnerText == null) ? defaultValue : destNode.Attributes[attribute].InnerText;
            }
            else { returnValue = defaultValue; }
            return returnValue;
        }

        public static XlRgbColor GetColor(string Color)
        {
            switch (Color.ToLower())
            {
                case "yellow":
                    return XlRgbColor.rgbYellow;
                case "gray":
                    return XlRgbColor.rgbGray;
                case "darkgray":
                    return XlRgbColor.rgbDarkGray;
                case "grey":
                    return XlRgbColor.rgbGrey;
                case "darkgrey":
                    return XlRgbColor.rgbDarkGrey;
                case "darkblue":
                    return XlRgbColor.rgbDarkBlue;
                case "lightblue":
                    return XlRgbColor.rgbLightBlue;
                case "navyblue":
                    return XlRgbColor.rgbNavyBlue;
                case "royalblue":
                    return XlRgbColor.rgbRoyalBlue;
                case "navy":
                    return XlRgbColor.rgbNavy;
                case "blue":
                    return XlRgbColor.rgbBlue;
                case "brown":
                    return XlRgbColor.rgbBrown;
                case "darkmagenta":
                    return XlRgbColor.rgbDarkMagenta;
                case "orange":
                    return XlRgbColor.rgbOrange;
                case "orangeRed":
                    return XlRgbColor.rgbOrangeRed;
                case "green":
                    return XlRgbColor.rgbGreen;
                case "darkgreen":
                    return XlRgbColor.rgbDarkGreen;
                case "seagreen":
                    return XlRgbColor.rgbSeaGreen;
                case "white":
                    return XlRgbColor.rgbWhite;
                case "red":
                    return XlRgbColor.rgbRed;
                case "gainsboro":
                    return XlRgbColor.rgbGainsboro;
                case "turquoise":
                    return XlRgbColor.rgbTurquoise;
                case "peachpuff":
                    return XlRgbColor.rgbPeachPuff;
                case "purple":
                    return XlRgbColor.rgbPurple;
                case "silver":
                    return XlRgbColor.rgbSilver;
                case "violet":
                    return XlRgbColor.rgbViolet;
                case "darkcyan":
                    return XlRgbColor.rgbDarkCyan;
                case "lightcyan":
                    return XlRgbColor.rgbLightCyan;
                case "lavender":
                    return XlRgbColor.rgbLavender;
                default:
                    return XlRgbColor.rgbBlack;

            }
        }


        #region Kill Process
        public static void KillSpecificProcess(string processName, string userName = "")
        {
            try
            {
                var processes = from p in Process.GetProcessesByName(processName)
                                select p;

                foreach (Process process in processes)
                {
                    try
                    {
                        if (!string.IsNullOrEmpty(userName) == true)
                        {
                            string ownerName = GetProcessOwner(process.Id);
                            if (userName.Equals(ownerName, StringComparison.InvariantCultureIgnoreCase))
                            {
                                process.Kill();
                                Logger.LogInfoMessage(string.Format("[KillSpecificProcess]" + " Process Name : " + processName + ", User Name : " + userName + " killed successfully "), true);
                            }
                        }
                        else
                        {
                            process.Kill();
                            Logger.LogInfoMessage(string.Format("[KillSpecificProcess]" + " Process Name : " + processName + " killed successfully "), true);
                        }
                    }
                    catch (Exception ex)
                    {
                        ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Pivot", ex.Message, ex.ToString(),
                            "KillSpecificProcess", ex.GetType().ToString(), Constants.NotApplicable);
                        Logger.LogErrorMessage(string.Format("[KillSpecificProcess][Exception]: " + ex.Message), true);
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Pivot", ex.Message, ex.ToString(),
                            "KillSpecificProcess", ex.GetType().ToString(), Constants.NotApplicable);
                Logger.LogErrorMessage(string.Format("[KillSpecificProcess][Exception]: " + ex.Message), true);
            }
        }
        public static string GetProcessOwner(int processId)
        {
            string query = "Select * From Win32_Process Where ProcessID = " + processId;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(query);
            ManagementObjectCollection processList = searcher.Get();

            foreach (ManagementObject obj in processList)
            {
                string[] argList = new string[] { string.Empty, string.Empty };
                int returnVal = Convert.ToInt32(obj.InvokeMethod("GetOwner", argList));
                if (returnVal == 0)
                {
                    // Domain Name will be in argList[1]
                    return argList[0];
                }
            }
            return "NO OWNER";
        }

        #endregion Kill Process

        #region isValidXML: Pivot Config XML Validation

        public static bool isValidXML(XmlNodeList Components)
        {
            bool isValid = true;

            //[START]Validation of PivotConfig XML
            if (Components.Count >= 1)
            {
                foreach (XmlNode component in Components)
                {
                    //Validating Component Attributes
                    if (component.Attributes.Count > 0)
                    {
                        isValid = isValidXML_Components(component, isValid);
                    }
                    else
                    {
                        Logger.LogInfoMessage(string.Format("[isValidXML] Config XML should contain all required Component's attributes. Ex. Name, InputFileName, InputDataSheetName, SummaryViewColumn"), true);
                        isValid = false;
                    }
                }
            }
            else
            {
                System.Console.WriteLine("Components count should be greater than 1");
                Logger.LogInfoMessage(string.Format("[isValidXML] Config XML should contain at leaset one <Component> Tag"), true);
                isValid = false;
            }
            //[END]Validation of PivotConfig XML

            return isValid;
        }
        public static bool isValidXML_Components(XmlNode component, bool isValid)
        {
            string componentName = string.Empty;
            string InputFileName = string.Empty;
            string inputDataSheetName = string.Empty;

            //[START]Validation of Component Tag 
            /*
             <Component Name="ContentType" InputFileName="ContentType_Usage.csv"
             *InputDataSheetName="ContentType_Usage" 
             *Description="Content Type Pivot and Slicers Creation">
             */
            if (component.Attributes["Name"] != null)
            {
                if (!string.IsNullOrEmpty(component.Attributes["Name"].InnerText))
                {
                    componentName = component.Attributes["Name"].InnerText;
                }
                else
                {
                    Logger.LogInfoMessage(string.Format("[isValidXML][isValidXML_Components] Component Attribute \"Name\" value should not be empty"), true);
                    isValid = false;
                }
            }
            else if (component.Attributes["Name"] == null)
            {
                Logger.LogInfoMessage(string.Format("[isValidXML][isValidXML_Components] Component Attribute \"Name\" should not be empty"), true);
                isValid = false;
            }

            if (component.Attributes["InputFileName"] != null)
            {
                if (!string.IsNullOrEmpty(component.Attributes["InputFileName"].InnerText))
                {
                    InputFileName = component.Attributes["InputFileName"].InnerText;
                }
                else
                {
                    Logger.LogInfoMessage(string.Format("[isValidXML][isValidXML_Components] Component Attribute \"InputFileName\" value should not be empty"), true);
                    isValid = false;
                }
            }
            else if (component.Attributes["InputFileName"] == null)
            {
                Logger.LogInfoMessage(string.Format("[isValidXML][isValidXML_Components] Component Attribute \"InputFileName\" should not be empty"), true);
                isValid = false;
            }

            if (component.Attributes["InputDataSheetName"] != null)
            {
                if (!string.IsNullOrEmpty(component.Attributes["InputDataSheetName"].InnerText))
                {
                    inputDataSheetName = component.Attributes["InputDataSheetName"].InnerText;
                }
                else
                {
                    System.Console.WriteLine("Component inputDataSheetNameshould not be empty for Component---" + componentName);
                    isValid = false;
                }
            }
            else if (component.Attributes["InputDataSheetName"] == null)
            {
                System.Console.WriteLine("Component inputDataSheetNameshould not be empty for Component---" + componentName);
                isValid = false;
            }
            //[END]Validation of Component Tag 

            //Calling Inner Validation Functions to Validate the Item Tag and it's Inner tags
            isValid = isValidXML_Item(component, componentName, isValid);

            return isValid;
        }
        public static bool isValidXML_Item(XmlNode component, string componentName, bool isValid)
        {
            //Validation for pivot and slicer view
            if (component.SelectNodes("Item").Count > 0)
            {
                foreach (XmlNode ItemNode in component.SelectNodes("Item"))
                {
                    if (ItemNode.Attributes.Count > 0)
                    {
                        if (ItemNode.Attributes["Type"] != null)
                        {
                            string itemType = ItemNode.Attributes["Type"].InnerText;
                            if (itemType == "PivotView")
                            {
                                //Calling Inner Validation Functions to Validate the "PivotView" Tag and it's Inner tags
                                isValid = isValidXML_PivotView(ItemNode, componentName, isValid);
                            }
                            else if (itemType == "SliceDiceView")
                            {
                                //Calling Inner Validation Functions to Validate the "SliceDiceView" Tag and it's Inner tags
                                isValid = isValidXML_SliceDiceView(ItemNode, componentName, isValid);
                            }
                            else
                            {
                                System.Console.WriteLine("Item type should be SliceDiceView or PivotView  for Component--" + componentName);
                                isValid = false;
                            }
                        }
                        else
                        {
                            System.Console.WriteLine("Type in Item Tag Should not empty for Component--" + componentName);
                            isValid = false;
                        }
                    }
                    else
                    {
                        System.Console.WriteLine("Attributes for Item Tag Should not empty for Component--" + componentName);
                        isValid = false;
                    }
                }
            }
            else
            {
                System.Console.WriteLine("Component Item tags are not presented for Components--" + componentName);
                isValid = false;
            }

            return isValid;
        }
        public static bool isValidXML_PivotView(XmlNode ItemNode, string componentName, bool isValid)
        {
            //[START]Validation of Value Tags
            //<ValueColumn Name="WebUrl" Function="Count"> </ValueColumn>

            if (ItemNode.SelectSingleNode("ValueColumn") != null)
            {
                var valueColumnField = ItemNode.SelectSingleNode("ValueColumn");
                if (valueColumnField.Attributes["Name"] != null)
                {
                    if (string.IsNullOrEmpty(valueColumnField.Attributes["Name"].InnerText))
                    {
                        System.Console.WriteLine("Value Coumn Name should not be empty for Componenet--" + componentName);
                        isValid = false;
                    }
                }
                else
                {
                    System.Console.WriteLine("Value Coumn Name is missing for Componenet--" + componentName);
                    isValid = false;
                }
            }
            else
            {
                System.Console.WriteLine("Value Column Tag is Missing for Component--" + componentName);
                isValid = false;
            }
            //[END]Validation of Value Tags

            //[START] Validation of SlicersStyling Tags
            //<SlicersStyling Style="SlicerStyleDark4" Height="" Width=""></SlicersStyling>
            if (ItemNode.SelectSingleNode("SlicersStyling") != null)
            {
                var slicerstyle = ItemNode.SelectSingleNode("SlicersStyling");
                if (slicerstyle.Attributes["Style"] != null)
                {
                    if (string.IsNullOrEmpty(slicerstyle.Attributes["Style"].InnerText))
                    {
                        System.Console.WriteLine("SlicersStyling Style is missing for Componenet--" + componentName);
                        isValid = false;
                    }
                }
                else
                {
                    System.Console.WriteLine("SlicersStyling Style is missing for Componenet--" + componentName);
                    isValid = false;
                }
            }
            else
            {
                System.Console.WriteLine("SlicersStyling Tag is Missing for Component--" + componentName);
                isValid = false;
            }
            //[END] Validation of SlicersStyling Tags

            //[START] Validation of Rows and Row Tags
            /*
             <Rows>
                <Row Column="ContentTypeName" Label="Content Type Name" Order="1"></Row>
             </Rows>
             */

            if (ItemNode.SelectNodes("Rows") != null)
            {
                if (ItemNode.SelectNodes("Rows").Count > 0)
                {
                    foreach (XmlNode RowFieldRoot in ItemNode.SelectNodes("Rows"))
                    {
                        if (RowFieldRoot.ChildNodes.Count > 0)
                        {
                            foreach (XmlNode rowFeild in RowFieldRoot.ChildNodes)
                            {
                                if (rowFeild.Attributes["Column"] != null)
                                {
                                    if (string.IsNullOrEmpty(rowFeild.Attributes["Column"].InnerText))
                                    {
                                        System.Console.WriteLine("Columns attribute should not be empty in Item type Pivot for Component--" + componentName);
                                        isValid = false;
                                    }
                                }
                                else
                                {
                                    System.Console.WriteLine("Columns attribute should be exist in Item type Pivot for Component--" + componentName);
                                    isValid = false;
                                }
                            }
                        }
                        else
                        {
                            System.Console.WriteLine("Atlease one Row should be exist in Item type Pivot for Component--" + componentName);
                            isValid = false;
                        }
                    }
                }
                else
                {
                    System.Console.WriteLine("Rows tag should be exist in Item type Pivot for Component--" + componentName);
                    isValid = false;
                }
            }
            else
            {
                System.Console.WriteLine("Rows tag should be exist in Item type Pivot for Component--" + componentName);
                isValid = false;
            }
            //[END] Validation of Rows and Row Tags

            //[START] Validation of Filters Tags
            /*
              <Filters>
                <Filter>WebApplicationUrl</Filter>
                <Filter>SiteCollectionUrl</Filter>
              </Filters>
             */
            if (ItemNode.SelectNodes("Filters") != null && ItemNode.SelectNodes("Filters").Count > 0)
            {

                foreach (XmlNode FilterFeildRoot in ItemNode.SelectNodes("Filters"))
                {
                    if (isValid && FilterFeildRoot.ChildNodes.Count == 0 || (FilterFeildRoot.ChildNodes.Count == 0))
                    {
                        System.Console.WriteLine("Atleast one filter should be exist in Item type Pivot for Component--" + componentName);
                        isValid = false;
                    }
                }
            }
            else
            {
                System.Console.WriteLine("Filters tag should be exist in Item type Pivot for Component--" + componentName);
                isValid = false;
            }
            //[END] Validation of Filters Tags

            //[START] Validation of Slicers Tags
            /*
              <Slicers>
                <Slicer>WebApplicationUrl</Slicer>
                <Slicer>SiteCollectionUrl</Slicer>
                <Slicer>isFromFeature</Slicer>
                <Slicer>ListTitle</Slicer>
                <Slicer>ListStatus</Slicer>
                <Slicer>SolutionName</Slicer>
              </Slicers>
             */
            if (ItemNode.SelectNodes("Slicers") != null && ItemNode.SelectNodes("Slicers").Count > 0)
            {
                foreach (XmlNode SlicersRoot in ItemNode.SelectNodes("Slicers"))
                {
                    if (isValid && SlicersRoot.ChildNodes.Count == 0)
                    {
                        System.Console.WriteLine("Atleast one slicer should be exist in Item type Pivot for Component--" + componentName);
                        isValid = false;
                    }
                }
            }
            else
            {
                System.Console.WriteLine("Slicers tag should be exist in Item type Pivot for Component--" + componentName);
                isValid = false;
            }
            //[END] Validation of Slicers Tags

            return isValid;
        }
        public static bool isValidXML_SliceDiceView(XmlNode ItemNode, string componentName, bool isValid)
        {

            //[START] Validation of Views Tags
            /*
             <Views>
                <View id="1">
                  <ValueColumn Name="WebApplicationUrl" Function="Count"> </ValueColumn>
                  <Rows>
                    <Row>ListStatus</Row>
                  </Rows>
                </View>
             </Views>
             */

            if (ItemNode.SelectNodes("Views") != null && ItemNode.SelectNodes("Views").Count > 0)
            {
                //Loop views for each Component
                foreach (XmlNode viewsNode in ItemNode.SelectNodes("Views"))
                {
                    if (viewsNode.ChildNodes != null)
                    {
                        XmlNode viewField = viewsNode.SelectSingleNode("View");
                        if (viewField != null)
                        {
                            var valuecoumnField = viewField.SelectSingleNode("ValueColumn");
                            if (valuecoumnField != null)
                            {
                                if (valuecoumnField.Attributes["Name"] != null)
                                {
                                    //[START] Validation of Rows Tags
                                    /*
                                     <Rows>
                                        <Row>ListStatus</Row>
                                     </Rows>       
                                     */
                                    if (viewField.SelectNodes("Rows").Count > 0)
                                    {
                                        foreach (XmlNode SlicersRoot in viewField.SelectNodes("Rows"))
                                        {
                                            if (isValid && SlicersRoot.ChildNodes.Count == 0)
                                            {
                                                System.Console.WriteLine("Atleast one row should be exist in Rows Tag in SliceDiceView type for Component--" + componentName);
                                                isValid = false;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        System.Console.WriteLine("Rows Tag should be in SliceDiceView type for Component--" + componentName);
                                        isValid = false;
                                    }
                                    //[END] Validation of Rows Tags
                                }
                            }
                            else
                            {
                                System.Console.WriteLine("ValueCoulmn name attribute is empty in Views Tag SliceDiceView type for Component--" + componentName);
                                isValid = false;
                            }
                        }
                    }
                    else
                    {
                        System.Console.WriteLine("ValueCoulmn tag is empty in Views Tag SliceDiceView type for Component--" + componentName);
                        isValid = false;
                    }
                }
            }
            else
            {
                System.Console.WriteLine("Views tag is empty in Itemtype SliceDiceView for Component--" + componentName);
                isValid = false;
            }

            //[END] Validation of Views Tags
            return isValid;
        }

        #endregion isValidXML: Pivot Config XML Validation
    }
}
