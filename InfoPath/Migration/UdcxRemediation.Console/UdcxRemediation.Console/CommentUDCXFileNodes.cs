using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UdcxRemediation.Console.Common.Base;
using UdcxRemediation.Console.Common.CSV;
using Microsoft.SharePoint.Client;
using System.Xml.Linq;
using System.IO;
using System.Xml;

namespace UdcxRemediation.Console
{
    public class CommentUDCXFileNodes
    {
        public static void DoWork(string inputFileSpec)
        {
            Logger.OpenLog("CommentUDCXFileNodes");
            Logger.LogInfoMessage(String.Format("Scan starting {0}", DateTime.Now.ToString()), true);
            Logger.LogInfoMessage(inputFileSpec, true);

            List<UdcxReportOutput> _WriteUDCList = null;
           
            IEnumerable<UdcxReportInput> udcxCSVRows = ImportCSV.ReadMatchingColumns<UdcxReportInput>(inputFileSpec, Constants.CsvDelimeter);
            if (udcxCSVRows != null)
            {
                try
                {
                    var authRows = udcxCSVRows.Where(x => x.Authentication != null && x.Authentication != Constants.ErrorStatus && x.Authentication.Length > 0);
                    if (authRows != null && authRows.Count() > 0)
                    {
                        _WriteUDCList = new List<UdcxReportOutput>();
                        Logger.LogInfoMessage(String.Format("Preparing to comment a total of {0} Udcx Files Nodes ...", authRows.Count()), true);

                        foreach (UdcxReportInput udcxFileInput in authRows)
                        {
                            CommentUDCXFileNode(udcxFileInput, _WriteUDCList);
                        }

                        GenerateStatusReport(_WriteUDCList);
                    }
                    else
                    {
                        Logger.LogInfoMessage("No valid authentication records found in '" + inputFileSpec + "' File ", true);
                    }                  
                }
                catch (Exception ex)
                {
                    Logger.LogErrorMessage(String.Format("CommentUDCXFileNode() failed: Error={0}", ex.Message), true);
                }

                Logger.LogInfoMessage(String.Format("Scan completed {0}", DateTime.Now.ToString()), true);
            }
            else
            {
                Logger.LogInfoMessage("No records found in '" + inputFileSpec + "' File ", true);
            }

            Logger.CloseLog();
        }

        private static void CommentUDCXFileNode(UdcxReportInput udcxFileInput, List<UdcxReportOutput> _WriteUDCList)
        {
            if (udcxFileInput == null)
            {
                return;
            }
                
            string siteUrl = udcxFileInput.SiteUrl;
            string webUrl = udcxFileInput.WebUrl;
            string dirName = udcxFileInput.DirName;
            string leafName = udcxFileInput.LeafName;
            string authentication = udcxFileInput.Authentication;
           
            UdcxReportOutput udcxOutput = new UdcxReportOutput();
            udcxOutput.SiteUrl = siteUrl;
            udcxOutput.WebUrl = webUrl;
            udcxOutput.DirName = dirName;
            udcxOutput.LeafName = leafName;
            udcxOutput.Authentication = authentication;

            if (dirName.EndsWith("/"))
            {
                dirName = dirName.TrimEnd(new char[] { '/' });
            }
            if (leafName.StartsWith("/"))
            {
                leafName = leafName.TrimStart(new char[] { '/' });
            }
            string serverRelativeFilePath = "/" + dirName + '/' + leafName;

            try
            {
                Logger.LogInfoMessage(String.Format("Processing UCDX File [{0}/{1}] of Web [{2}] ...", dirName, leafName, webUrl), true);
               
                using (ClientContext userContext = Helper.CreateAuthenticatedUserContext(Program.AdminDomain, Program.AdminUsername, Program.AdminPassword, siteUrl))
                {
                    userContext.ExecuteQuery();

                    FileInformation info = Microsoft.SharePoint.Client.File.OpenBinaryDirect(userContext, serverRelativeFilePath);

                    XNamespace xns = "http://schemas.microsoft.com/office/infopath/2006/udc";
                    XDocument doc = XDocument.Load(XmlReader.Create(info.Stream));                    
                    XElement authElem = doc.Root.Element(xns + "ConnectionInfo").Element(xns + "Authentication");

                    if (authElem != null)
                    {
                        string authData = authElem.ToString();
                        authData = authData.Replace("<udc:Authentication xmlns:udc=\"" + xns + "\">", "<udc:Authentication>");
                        authElem.ReplaceWith(new XComment(authData));
                        
                        string saveUdcxContent = doc.Declaration.ToString() + doc.ToString();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            StreamWriter writer = new StreamWriter(memoryStream);
                            writer.Write(saveUdcxContent);
                            writer.Flush();
                            memoryStream.Position = 0;

                            Microsoft.SharePoint.Client.File.SaveBinaryDirect(userContext, serverRelativeFilePath, memoryStream, true);
                        }

                        udcxOutput.Status = Constants.SuccessStatus;
                    }
                    else
                    {
                        udcxOutput.Status = Constants.NoAuthNodeFound;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogErrorMessage(String.Format("CommentUDCXFileNode() failed for UDCX File [{0}/{1}] of Web [{2}]: Error={3}", dirName, leafName, webUrl, ex.Message), false);
                udcxOutput.Status = Constants.ErrorStatus;
                udcxOutput.ErrorDetails = ex.Message;
            }
                        
            _WriteUDCList.Add(udcxOutput);            
        }

        private static void GenerateStatusReport(List<UdcxReportOutput> _WriteUDCList)
        {
            string reportFileName = Environment.CurrentDirectory + "\\" + Constants.UdcxReport;

            Common.Utilities.FileUtility.WriteCsVintoFile(reportFileName, ref _WriteUDCList);
        }

    }
}
