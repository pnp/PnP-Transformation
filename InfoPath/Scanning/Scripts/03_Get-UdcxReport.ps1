<#
.SYNOPSIS
Parses downloaded UDCX files and generates UDCX report.

.DESCRIPTION
Execute the script to read xml data of UDCX files. Retrieve 'Authentication, SelectCommand and UpdateCommand' nodes data and generates report.

.Example
.\03_Get-UdcxReport.ps1

#>

$getInfoPathFiles = {
    try
    {
        $ErrorActionPreference = "STOP"
        $retVal = $false

        $localPath = "c:\Scanner\InfoPath"   

        $udcxFileExtension = "udcx"
        $udcxFolder = Join-Path $localPath $udcxFileExtension
        $udcxInputFile = Join-Path $udcxFolder "InfoPathFileInfo.log"
        $reportFilePath = Join-Path $udcxFolder "UDCXReport.csv"
        $udcxTempFilePath = Join-Path $udcxFolder "{0}\{1}"

        $fileDateTime = [DateTime]::Now.ToString("MMddyyyy_hhmmss")  #Ensure the Output and Error Log Files have the same Date/Time name
        $errorFile = [string]::Format("{0}\UDCXReport_Errors_{1}.csv", $udcxFolder, $fileDateTime)
    
        function Error-To-File($errorMsg, $output)
        {
           $output | Add-Member -MemberType NoteProperty -Name Error -Value $errorMsg
           $output | Export-Csv -notypeinformation -Path $errorFile  -Append
        }

        $rowObjects = @()
        $csvRows=@(Import-Csv $udcxInputFile | sort SiteId,WebId,ID –Unique)          
        $i = 1;  

        foreach($csvRow in $csvRows)
        {
            try
            {
                $udcxFilePath = [string]::Format($udcxTempFilePath, $csvRow.Id, $csvRow.LeafName)
                $relativeUrl = [string]::Format("{0}/{1}", $csvRow.DirName, $csvRow.LeafName)

                $csvRow | Add-Member RelativeUrl $relativeUrl
                
                if(Test-Path $udcxFilePath)
                {
                    [xml]$xmlContent = Get-Content $udcxFilePath
                    $ns = New-Object System.Xml.XmlNamespaceManager($xmlContent.NameTable)
                    $ns.AddNamespace("udc", "http://schemas.microsoft.com/office/infopath/2006/udc")

                    $authNode=$xmlContent.SelectSingleNode("//udc:Authentication", $ns).InnerXML -replace '[\r\n]',''
                    $csvRow | Add-Member Authentication $authNode
                    
                    $selectNode = $xmlContent.SelectSingleNode("//udc:SelectCommand", $ns)
                    $csvRow | Add-Member SelectServiceUrl $selectNode.ServiceUrl.InnerXML

                    $selectSoapAction = $selectNode.SoapAction
                    $csvRow | Add-Member SelectSoapAction $selectSoapAction

                    $selectSoapActionName = ""
                    if ($selectSoapAction)
                    { 
                        $selectSoapActionName = $selectSoapAction.SubString($selectSoapAction.LastIndexOf("/")+1)
                    }

                    $csvRow | Add-Member SelectSoapActionName $selectSoapActionName
                    $queryNode=$selectNode.Query -replace '[\r\n]',''              
                    $csvRow | Add-Member SelectQuery $($queryNode)
                    
                    $updateNode = $xmlContent.SelectSingleNode("//udc:UpdateCommand", $ns)
                    $csvRow | Add-Member UpdateServiceUrl $updateNode.ServiceUrl.InnerXML

                    $updateSoapAction = $updateNode.SoapAction
                    $csvRow | Add-Member UpdateSoapAction $updateSoapAction

                    $updateSoapActionName = ""
                    if ($updateSoapAction)
                    { 
                        $updateSoapActionName = $updateSoapAction.SubString($updateSoapAction.LastIndexOf("/")+1)
                    }

                    $csvRow | Add-Member UpdateSoapActionName $updateSoapActionName     
                                  
                    $rowObjects += $csvRow
                   
                }
                else
                {
                    Error-To-File "File Not Found $($udcxFilePath)" $csvRow

                    $csvRow | Add-Member Authentication "Error"
                    $csvRow | Add-Member SelectServiceUrl "Error"
                    $csvRow | Add-Member SelectSoapAction "Error"
                    $csvRow | Add-Member SelectSoapActionName "Error"
                    $csvRow | Add-Member SelectQuery "Error"
                    $csvRow | Add-Member UpdateServiceUrl "Error"
                    $csvRow | Add-Member UpdateSoapAction "Error"
                    $csvRow | Add-Member UpdateSoapActionName "Error" 
                    $rowObjects += $csvRow
                }
            }
            catch [Exception]
            {
                Error-To-File $_.Exception.ToString()  $csvRow

                $csvRow | Add-Member Authentication "Error"
                $csvRow | Add-Member SelectServiceUrl "Error"
                $csvRow | Add-Member SelectSoapAction "Error"
                $csvRow | Add-Member SelectSoapActionName "Error"
                $csvRow | Add-Member SelectQuery "Error"
                $csvRow | Add-Member UpdateServiceUrl "Error"
                $csvRow | Add-Member UpdateSoapAction "Error"
                $csvRow | Add-Member UpdateSoapActionName "Error" 
                $rowObjects += $csvRow
            }
        }

        $rowObjects | Export-CSV -notypeinformation $reportFilePath
        
        Write-Host "Completed"
        $retVal = $true
    }
    catch [Exception]
    {
        $retVal = $_.CategoryInfo.Activity + " : " + $_.Exception.Message + $_.InvocationInfo.PositionMessage
        $retVal = $retVal.Replace("`"", "`'")
        $retVal = $retVal.Replace("'", "``'")
    }

    return $retVal
}

return (Powershell.exe -command $getInfoPathFiles)