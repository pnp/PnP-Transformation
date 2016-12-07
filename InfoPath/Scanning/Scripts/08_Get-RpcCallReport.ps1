<#
.SYNOPSIS
Parses the output from 02_Scrape-InfoPathFiles.ps1 to generate a customer friendly report.

The report identifies those InfoPath form (XSN) instances that use a UDCX file to call 
the RPC endpoint (OWSSVR.dll).  RPC calls will fail in SPO-MT/vNext due to loopback protection; 
as such, the offending forms/udcx need to be remediated.

TBD: The report needs to be modified to consider embedded RPC connections once we figure out how they are represented in the XSN.

.DESCRIPTION
Execute the script to generate a report of RPC (owssvr.dll) Calls

.Example
.\0x_Get-RpcCallReport.ps1

#>
try
{
    #======================================================================================
    #ACTION ITEMS: Please initialize these settings with values that suit your environment

    #specify the location of your InfoPath Report folder
    $Global:exportFolder = "C:\scanner\InfoPath"

    # list all of your supported managed paths here
    # - do not include any slash ("/") characters
    $Global:managedPaths = @("sites", "personal")
    #=======================================================================

    $retVal = $false

    $ErrorActionPreference = "STOP"

    $Global:reportFolder = "$Global:exportFolder\xsn"
    $Global:udcxFolder = "$Global:exportFolder\udcx"

    #these collections are populated when we parse the InfoPathScraper_report.csv file
    $Global:cabPaths = @()
    $Global:managedCode = @()
    $Global:modes = @()
    $Global:productVersions = @()
    $Global:publishUrls = @()
    $Global:soapConnections = @()
    $Global:restConnections = @()
    $Global:udcxConnections = @()
    $Global:adoConnections = @()
    $Global:dataConnections = @()
    $Global:xmlConnections = @()
    $Global:urns = @()

    #Each member of this dictionary represents a row in the output report
    $reportOutput = @{}
    #each member of this dictionary represents a row from the InfoPathFileInfo.log
    $fileInfo = @{}

    $fileDateTime = [DateTime]::Now.ToString("MMddyyyy_hhmmss")  #Ensure the Output and Error Log Files have the same Date/Time name

    $outputFile = [string]::Format("InfoPathForms_RpcCall_{0}.csv", $fileDateTime)
    $errorFile = [string]::Format("InfoPathForms_RpcCall_Errors_{0}.csv", $fileDateTime)

    if (-not (Test-Path $Global:exportFolder)) 
    { 
        New-Item $Global:exportFolder -ItemType Directory | Out-Null
    }

    #This is the Form Template (XSN) contents/attributes report
    $infoPathRecords = Import-CSV $Global:reportFolder\InfoPathScraper_report.csv -Header ColumnA, ColumnB, ColumnC, ColumnD
    
    #This is the UDCX File instance report
    $udcxRecords = @(Import-CSV $Global:udcxFolder\UDCXReport.csv | Select WebId, SiteId, WebUrl, SiteUrl, RelativeUrl, SelectQuery)
    
    #=========================================================================================
    #functions
    #=========================================================================================

    function Add-Urn
    {
        param([string] $urn)

        if(-not $Global:urns.Contains($urn))
        {
            $Global:urns += $urn
        }
    }

    function DeDupArray
    {
        param($array)

        return $array | select -Unique
    }

    function Get-XsnMode
    {
        param([PSObject]$currentObject)

        $ourModes = @($Global:modes | ?{$_.ColumnA -eq $currentObject.XsnUrn})

        foreach($mode in $ourModes)
        {
            if($mode.ColumnA -ne $currentObject.XsnUrn)
            {
                continue
            }    

            $currentObject.XsnMode = $mode.ColumnC

            break
        }
    }

    # returns a collection of records where each record describes a deployed instance of the specified XSN Form template
    function Get-XsnLocations
    {
        param($urn)

        $formCabPaths = @($Global:cabPaths | ?{$_.ColumnA -eq $urn})
    
        $retVal = @()

        foreach($url in $formCabPaths)
        {
            $tempCabPath = ""
            if([string]::IsNullOrWhiteSpace($url.ColumnC))
            {
                $tempCabPath = $url.ColumnD
            }
            else
            {
                $tempCabPath = $url.ColumnC
            }        
        
            #CabPath will always be in the following format c:\splocalbin\infopath_scanner\[Guid]\file.xsn       
            $docID = $tempCabPath.Substring(($tempCabPath.LastIndexOf("\")-36), 36)
            $file = $fileInfo[$docID]

            $tempVal = New-Object PSObject -Property @{ 
                XsnFileUrl = [string]::Format("{0}/{1}", $file.DirName, $file.LeafName)
                XsnSiteId = $file.SiteId
                XsnSiteUrl = $file.SiteUrl
                XsnWebId = $file.WebId
                XsnWebUrl = $file.WebUrl
                }

            $retVal += $tempVal
        }

        return $retVal
    }

    # loads information for all of RPC (owssvr.dll) service calls issued by all deployed instances of the specified XSN form template
    # examines both embedded and UDCX-based RPC calls
    #
    function Get-RpcConnections
    {
        param([PSObject]$currentObject)  # the XSN form template to process

        $xsnSiteUrls = @()    
        $xsnFileUrls = @()    
        $udcxSiteUrls = @()    
        $udcxFileUrls = @()    
        $udcxRpcUrls = @()
        $xmlRpcUrls = @()

        #construct a query string that will locate RPC calls
        $rpcCallPath = "*owssvr.dll*"
                
        #query the UDCX connection definitions read from the scraper report for the ones defined for this XSN form template
        $ourUdcxConnections = @($Global:udcxConnections |
            ? {$_.ColumnA -eq $currentObject.XsnUrn} |
            Select -Unique ColumnA, ColumnB, ColumnC, ColumnD)

        #query the Xml connection definitions read from the scraper report for the ones defined for this XSN form template
        $ourXmlConnections = @($Global:xmlConnections |
            ? {$_.ColumnA -eq $currentObject.XsnUrn -and $_.ColumnC -like $rpcCallPath} |
            Select -Unique ColumnA, ColumnB, ColumnC, ColumnD)

        $hasConnectionToProcess = ($ourUdcxConnections.Count -gt 0) -or ($ourXmlConnections.Count -gt 0)            
        if($hasConnectionToProcess)
        {            
            #get the list of locations where this XSN form template has been deployed
            $xsnLocations = Get-XsnLocations $currentObject.XsnUrn

            #for each deployed instance of this form template...
            foreach ($xsnLocation in $xsnLocations)
            {
                #for each udcx connection defined for this XSN form template...
                foreach($item in $ourUdcxConnections)
                {
                    #---------------------------------------------------------
                    # Attempt to locate the expected UDCX File instance...
                    #---------------------------------------------------------

                    #construct a query string to locate UDCX connection files whose name matches the SPSite-relative path of the expected UDCX connection file
                    $udcxFilePath = [string]::Format("*{0}", $item.ColumnC)   

                    #Query the list of UDCX Connection File records for the one that is _expected_ by this deployed form instance
                    # Notes: 
                    #  - XSN Forms can be deployed to a different SPWeb than the one that contains the UDCX file; however, both SPWebs must be within the same SPSite (so we compare SiteIds)
                    #  - It is entirely possible that the expected UDCX file is not present (e.g., not deployed, moved, renamed, deleted, etc.)
                    #  - The query might return more than one UDCX file instance; however, no harm done - net result is redundant report rows, not additional work to be done
                    #
                    $udcxFiles = @($udcxRecords | ? { $_.SiteId -eq $xsnLocation.XsnSiteId -and $_.SelectQuery -like $rpcCallPath -and $_.RelativeUrl -like  $udcxFilePath})

                    foreach($udcxFile in $udcxFiles)
                    {
                        #success: we found the UDCX file instance... 
                        # add the information for this udcx connection instance to the supporting collections of this form template record
                        # be sure to keep the collections in sync                        

                        #add the location of the form instance
                        $xsnSiteUrls += $xsnLocation.XsnSiteUrl.ToLower()
                        $xsnFileUrls += $xsnLocation.XsnFileUrl.ToLower()

                        #add the location of the udcx file referenced
                        $udcxSiteUrls += $udcxFile.SiteUrl.ToLower()
                        $udcxFileUrls += $udcxFile.RelativeUrl.ToLower()

                        #add the RPC Call Url contained in the referenced udcx file
                        $udcxRpcUrls += $udcxFile.SelectQuery.ToLower()

                        #add a blank entry to keep the collections in sync
                        $xmlRpcUrls += ""
                    }
                }

                #for each embedded connection defined for this XSN form template...
                foreach($item in $ourXmlConnections)
                {
                    #---------------------------------------------------------
                    # Attempt to locate the expected RPC Call...
                    #---------------------------------------------------------
                    if ($item.ColumnC -like $rpcCallPath)
                    {
                        #success: we found an Xml Rpc Call... 
                        # add the information for this connection instance to the supporting collections of this form template record
                        # be sure to keep the collections in sync                        

                        #add the location of the form instance
                        $xsnSiteUrls += $xsnLocation.XsnSiteUrl.ToLower()
                        $xsnFileUrls += $xsnLocation.XsnFileUrl.ToLower()

                        #add a blank entry to keep the collections in sync
                        $udcxSiteUrls += ""
                        $udcxFileUrls += ""

                        #add a blank entry to keep the collections in sync
                        $udcxRpcUrls += ""

                        #add the RPC Call Url contained in the form template
                        $xmlRpcUrls += $item.ColumnC.ToLower()
                    }
                }
            }
        }

        # Assert: $xsnFileUrls.Count -eq $udcxFileUrls.Count -eq $udcxSupportedServiceSiteUrls.Count
        #
        $currentObject.XsnInstanceCount = $xsnFileUrls.Count
        if ($xsnFileUrls.Count -gt 0)
        {
            # Assert: $xsnSiteUrls.Count -eq $xsnFileUrls.Count
            $currentObject.XsnSiteUrls = [string]::Join(";", ($xsnSiteUrls)).ToLower() -replace '[\r\n]',''
            $currentObject.XsnFileUrls = [string]::Join(";", ($xsnFileUrls)).ToLower() -replace '[\r\n]',''
        }

        $currentObject.UdcxFileUrlsCount = $udcxFileUrls.Count
        if ($udcxFileUrls.Count -gt 0)
        {
            # Assert: $udcxSiteUrls.Count -eq $udcxFileUrls.Count
            $currentObject.UdcxSiteUrls = [string]::Join(";", ($udcxSiteUrls)).ToLower() -replace '[\r\n]',''
            $currentObject.UdcxFileUrls = [string]::Join(";", ($udcxFileUrls)).ToLower() -replace '[\r\n]',''
        }
        
        $currentObject.UdcxRpcUrlsCount = $udcxRpcUrls.Count
        if ($udcxRpcUrls.Count -gt 0)
        {
            $currentObject.UdcxRpcUrls = [string]::Join(";", ($udcxRpcUrls)).ToLower() -replace '[\r\n]',''
        }

        $currentObject.XmlRpcUrlsCount = $xmlRpcUrls.Count
        if ($xmlRpcUrls.Count -gt 0)
        {
            $currentObject.XmlRpcUrls = [string]::Join(";", ($xmlRpcUrls)).ToLower() -replace '[\r\n]',''
        }
    }    


    function Get-Remediation
    {
        param([PSObject]$currentObject)

        #The only option is to remove the RPC Call
        $currentObject.Remediation = "Remove RPC"
        return
    }

    function Process-Collections
    {
        foreach($urn in $Global:urns)
        {

            #this represents the current XSN form template.
            $currentObject = New-Object PSObject -Property @{ 
                XsnUrn = $urn
                XsnMode = [string]::Empty
                XsnInstanceCount = 0
                XsnSiteUrls = [string]::Empty
                XsnFileUrls = [string]::Empty
                UdcxFileUrlsCount = 0
                UdcxSiteUrls = [string]::Empty
                UdcxFileUrls = [string]::Empty
                UdcxRpcUrlsCount = 0                                                              
                UdcxRpcUrls = [string]::Empty
                XmlRpcUrlsCount = 0                                                              
                XmlRpcUrls = [string]::Empty
                Remediation = [string]::Empty
                }

            #load information for all of the RPC-based service calls issued by all deployed instances of this XSN form template
            Get-RpcConnections ([REF]$currentObject)
            
            $hasRpcCall = ($currentObject.UdcxRpcUrlsCount -gt 0) -or ($currentObject.XmlRpcUrlsCount -gt 0)
            
            if($hasRpcCall)
            {
                Get-XsnMode ([REF]$currentObject)

                # We will now output a row for each UDCX-based Service call issued by all deployed instances of this XSN form template

                # We effectively have several sets of parallel collections for this XSN form instance:
                # - for a given index value (e.g., 'n'), the set of items at that index across each collection represents a single RPC call
                #
                $local:xsnSiteUrls = $currentObject.XsnSiteUrls -split ";"
                $local:xsnFileUrls = $currentObject.XsnFileUrls -split ";"
                $local:udcxSiteUrls = $currentObject.UdcxSiteUrls -split ";"
                $local:udcxFileUrls = $currentObject.UdcxFileUrls -split ";"
                $local:udcxRpcUrls = $currentObject.UdcxRpcUrls -split ";"
                $local:xmlRpcUrls = $currentObject.XmlRpcUrls -split ";"

                #expand/denormalize the parallel collections of each row
                #
                for ($i=0; $i -lt $currentObject.XsnInstanceCount; $i++)
                {
                    $tempObject = New-Object PSObject -Property @{ 
                        XsnUrn = $currentObject.XsnUrn
                        XsnMode = $currentObject.XsnMode
                        XsnInstanceCount = $currentObject.XsnInstanceCount
                        XsnSiteUrls = [string]::Empty
                        XsnFileUrls = [string]::Empty
                        UdcxFileUrlsCount = $currentObject.UdcxFileUrlsCount
                        UdcxSiteUrls = [string]::Empty
                        UdcxFileUrls = [string]::Empty
                        UdcxRpcUrlsCount = $currentObject.UdcxRpcUrlsCount                                                       
                        UdcxRpcUrls = [string]::Empty
                        XmlRpcUrlsCount = $currentObject.XmlRpcUrlsCount 
                        XmlRpcUrls = [string]::Empty
                        Remediation = [string]::Empty
                        }

                    $tempObject.XsnSiteUrls = $local:xsnSiteUrls[$i]
                    $tempObject.XsnFileUrls = $local:xsnFileUrls[$i]

                    $tempObject.UdcxSiteUrls = $local:udcxSiteUrls[$i]
                    $tempObject.UdcxFileUrls = $local:udcxFileUrls[$i]

                    $tempObject.UdcxRpcUrls = $local:udcxRpcUrls[$i]
                    $tempObject.XmlRpcUrls = $local:xmlRpcUrls[$i]

                    # Now that the row has been normalized, we can compute the remediation for the form/udcx represented by the row
                    Get-Remediation ([REF]$tempObject)

                    #Add the row to the output collection.  Since we add multiple report rows for this URN, we need to generate a unique key for each Add operation related to this form
                    # Note:
                    # - $urn comes from the $Global:urns array and the Contains operator used in Add-Urn is case-insensitive.
                    # - $reportOutput is a dictionary and its keys are case-insensitive as well.
                    #
                    # Therefore, $Global:urns might contain two real InfoPath Form Template urns that differ only by case; each represents a unique form template in the eyes of InfoPath
                    # - urn:schemas-microsoft-com:office:infopath:sampleform:-myXSD-2015-03-03T20-34-02
                    # - urn:schemas-microsoft-com:office:infopath:SAMPLEFORM:-myXSD-2015-03-03T20-34-02
                    #
                    # In this case, we will process the first one and add its multiple report rows to $reportOutput without issue
                    # When we try to later process the second urn, we will get a key collision on the Add operation if we use a suffix format such as (e.g., "#" + i.ToString())
                    # We now use a suffix format that is practically guaranteed to be unique
                    $ticks = [DateTime]::Now.Ticks
                    $key = $urn + '#' + $i.ToString() + $ticks.ToString();
                    $reportOutput.Add($key, $tempObject);
                }
            }        
        }
    }

    function Populate-Collections
    {
        #Read in the file template (XSN) location report
        foreach($row in Import-CSV $Global:reportFolder\InfoPathFileInfo.log)
        {
            if(-not $fileInfo.ContainsKey($row.Id))
            {
                $fileInfo.Add($row.Id, $row)    
            }
        }

        #construct a query string that will locate RPC calls
        $rpcCallPath = "*owssvr.dll*"
                
        #Read in the XSN contents/attributes report
        foreach($item in $infoPathRecords)
        {
            switch($item.ColumnB)
            {
                "PublishUrl" { $Global:publishUrls += $item }

                "CabPath" { $Global:cabPaths += $item }

                "Mode" { $Global:modes += $item }

                "UdcConnection" {
                        $Global:udcxConnections += $item
                        Add-Urn($item.ColumnA)
                    }

                "XmlConnection" { 
                        # We care only about RPC XmlConnections
                        if ($item.ColumnC -like $rpcCallPath)
                        {
                            $Global:xmlConnections += $item
                            Add-Urn($item.ColumnA)
                        }
                    }
            }
        }
    }


    function Generate-Report
    {
        #Load the input files and prepare the supporting data structures 
        Populate-Collections    

        #Process all supporting data structures
        Process-Collections    

        #Create the output report file
        if($reportOutput.Count -gt 0)
        {
            #We can remove duplicate rows now that all collections have been normalized
            $reportOutput.Values | 
                ?{$_.XsnMode.ToLower() -ne "client"} | 
                Select -Unique Remediation, XsnUrn, XsnInstanceCount, XsnSiteUrls, XsnFileUrls, UdcxSiteUrls, UdcxFileUrls, UdcxRpcUrls, XmlRpcUrls | 
                Sort-Object -CaseSensitive -Property XsnUrn, XsnSiteUrls, XsnFileUrls, UdcxSiteUrls, UdcxFileUrls, UdcxRpcUrls, XmlRpcUrls |
                Export-CSV "$Global:exportFolder\$outputFile" -NoTypeInformation 
        }
        else
        {
             "No impacted InfoPath forms were located." | Out-File "$Global:exportFolder\$outputFile"   
        }
    }


    #=====================================================
    # main logic
    #=====================================================

    #here we go...

    #Generate the report...
    Generate-Report

    $retVal = $true
}
catch [Exception]
{
    $retVal = $_.CategoryInfo.Activity + " : " + $_.Exception.Message + $_.InvocationInfo.PositionMessage
    $retVal = $retVal.Replace("`"", "`'")
    $retVal = $retVal.Replace("'", "``'")
}

return $retVal