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
using JDP.Remediation.Console.Common.CSV;
using JDP.Remediation.Console.Common.Base;
using JDP.Remediation.Console.Common.Utilities;


namespace JDP.Remediation.Console.PivotHelper
{
    class ChartHelper
    {
        /// <summary>
        /// Draws Chart or Graph for given component.
        /// </summary>
        /// <param name="componentName"></param>
        /// <param name="oSummarySheet"></param>
        /// <param name="webAppUrlColumn"></param>
        /// <param name="componentColumnCount"></param>
        /// <param name="counter"></param>
        /// <param name="row"></param>
        /// <param name="chartType"></param>
        /// <param name="chartWidth"></param>
        /// <param name="chartHeight"></param>
        /// <param name="chartStyle"></param>
        /// <param name="CellIndex"></param>
        public static void DrawGraph(string componentName, Excel.Worksheet oSummarySheet, string webAppUrlColumn, int componentColumnCount, int counter,
            int row, string chartType, int chartWidth, int chartHeight, int chartStyle, char CellIndex)
        {
            //Create chart object
            Excel.Shape _Shape = oSummarySheet.Shapes.AddChart2();

            //Specify type of chart
            if (chartType.Equals("pie"))
                _Shape.Chart.ChartType = Excel.XlChartType.xlPie;
            else if (chartType.Equals("3dpie"))
                _Shape.Chart.ChartType = Excel.XlChartType.xl3DPie;
            else if (chartType.Equals("line"))
                _Shape.Chart.ChartType = Excel.XlChartType.xlLine;
            else if (chartType.Equals("3dline"))
                _Shape.Chart.ChartType = Excel.XlChartType.xl3DLine;
            else if (chartType.Equals("3dcolumn"))
                _Shape.Chart.ChartType = Excel.XlChartType.xl3DColumn;
            else if (chartType.Equals("clusteredcolumn"))
                _Shape.Chart.ChartType = Excel.XlChartType.xlColumnClustered;
            else if (chartType.Equals("3dclusteredcolumn"))
                _Shape.Chart.ChartType = Excel.XlChartType.xl3DColumnClustered;

            //Series object for the graph
            Excel.Series series = null;
            string exceptionComment = "[DrawGraph] Processing for Component :" + componentName;
            Logger.LogInfoMessage(String.Format("[GeneratePivotReports][DrawGraph] Processing Started for (" + componentName + ")"), false);

            try
            {
                //Get Series Column from SummaryView table
                string componentColumn = ((Char)(Convert.ToUInt16(CellIndex) + componentColumnCount)).ToString();

                //Set Series column for the Graph
                series = _Shape.Chart.SeriesCollection().Add(oSummarySheet.Range[componentColumn + row + ":"
                    + componentColumn + (counter).ToString()]);
                //Set Categories column for the graph
                series.XValues = oSummarySheet.Range[webAppUrlColumn + (row + 1).ToString() + ":"
                    + webAppUrlColumn + (counter).ToString()];

                //Apply data labels for the graph
                _Shape.Chart.ApplyDataLabels(Excel.XlDataLabelsType.xlDataLabelsShowBubbleSizes);
                //Apply legend for the graph
                _Shape.Chart.HasLegend = true;
                //apply style to chart
                _Shape.Chart.ChartStyle = chartStyle;

                //Hide Display Labels when their value is zero (0)
                Excel.SeriesCollection oSeriesCollection = (Excel.SeriesCollection)_Shape.Chart.SeriesCollection(Type.Missing);
                for (int j = 1; j <= oSeriesCollection.Count; j++)
                {
                    Excel.Series oSeries = (Excel.Series)oSeriesCollection.Item(j);
                    System.Array Values = (System.Array)((object)oSeries.Values);
                    //Array Values = (Array)oSeries.Values;
                    for (int k = 1; k <= Values.Length; k++)
                    {
                        Excel.DataLabel oDataLabel = (Excel.DataLabel)oSeries.DataLabels(k);
                        string caption = oDataLabel.Caption.ToString();
                        if (caption.Equals("0"))
                        {
                            oDataLabel.ShowValue = false;
                        }
                    }
                }

                //Set the Size of the Chart 
                _Shape.Width = chartWidth;
                _Shape.Height = chartHeight;

                //Calculations for the position of Chart
                int columnIndex = counter + 3;
                if (componentColumnCount > 3 && (componentColumnCount / 3 > 0))
                    columnIndex = (counter + 3) + (((componentColumnCount - 1) / 3) * 16);

                string charPositionColumn = webAppUrlColumn;

                if (componentColumnCount % 3 == 2)
                {
                    charPositionColumn = ((Char)(Convert.ToUInt16(CellIndex) + 7)).ToString();
                }

                if (componentColumnCount % 3 == 0)
                {
                    charPositionColumn = ((Char)(Convert.ToUInt16(CellIndex) + 14)).ToString();
                }

                _Shape.Left = (float)oSummarySheet.get_Range(charPositionColumn + columnIndex.ToString()).Left;
                _Shape.Top = (float)oSummarySheet.get_Range(charPositionColumn + columnIndex.ToString()).Top;

                Logger.LogInfoMessage(String.Format("[GeneratePivotReports][DrawGraph] Process Completed for (" + componentName + ")"), true);
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("[GeneratePivotReports][DrawGraph][Exception]: " + ex.Message + ", " + exceptionComment), true);
                ExceptionCsv.WriteException(Constants.NotApplicable, Constants.NotApplicable, Constants.NotApplicable, "Pivot", ex.Message, ex.ToString(),
                            "[GeneratePivotReports]: DrawGraph", ex.GetType().ToString(), exceptionComment);
            }
            finally
            {
                _Shape = null;
                series = null;
                oSummarySheet = null;
            }
        }
    }
}
