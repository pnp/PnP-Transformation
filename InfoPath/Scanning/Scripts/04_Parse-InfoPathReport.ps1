<#
.SYNOPSIS
Parses the output from 02_Scrape-InfoPathFiles.ps1 to generate a customer friendly report.

.DESCRIPTION
Execute the script to parse out the report.csv file

.Example
.\04_Parse-InfoPathReport.ps1

#>
try
{
    $retVal = $false

    $ErrorActionPreference = "STOP"

    $Global:exportFolder = "c:\Scanner\InfoPath"

    $Global:reportFolder = "$Global:exportFolder\xsn"
    $Global:udcxFolder = "$Global:exportFolder\udcx"
    $Global:cabPaths = @()
    $Global:managedCode = @()
    $Global:modes = @()
    $Global:productVersions = @()
    $Global:publishUrls = @()
    $Global:soapConnections = @()
    $Global:restConnections = @()
    $Global:udcConnections = @()
    $Global:adoConnections = @()
    $Global:dataConnections = @()
    $Global:urns = @()

    $reportOutput = @{}
    $fileInfo = @{}

    $supportedWebServices = @{
                                "GetUserProfileByName" = "userprofileservice.asmx"
                                "GetGroupCollectionFromWeb" = "usergroup.asmx"	
                                "GetUserCollectionFromGroup" = "usergroup.asmx"
                                "GetCommonManager" = "userprofileservice.asmx"
                                "GetCommonMemberships" = "userprofileservice.asmx"
                                "GetUserMemberships" = "userprofileservice.asmx"
                                #"GetRunningWorkflowTasksForCurrentUser" = "workflow.asmx"
                                "GetUserPropertyByAccountName" = "userprofileservice.asmx"
                                "CheckInFile" = "lists.asmx"
                                "CheckOutFile" = "lists.asmx"
                                "GetUserCollectionFromSite" = "usergroup.asmx"
                            }

    $fileDateTime = [DateTime]::Now.ToString("MMddyyyy_hhmmss")  #Ensure the Output and Error Log Files have the same Date/Time name

    $outputFile = [string]::Format("InfoPathForms_{0}.csv", $fileDateTime)

    $errorFile = [string]::Format("InfoPathForms_Errors_{0}.csv", $fileDateTime)


    if (-not (Test-Path $Global:exportFolder)) 
    { 
        New-Item $Global:exportFolder -ItemType Directory | Out-Null
    }


    $infoPathRecords = Import-CSV $Global:reportFolder\InfoPathScraper_report.csv -Header ColumnA, ColumnB, ColumnC, ColumnD
    $udcxRecords = @(Import-CSV $Global:udcxFolder\UDCXReport.csv | Select SiteId, RelativeUrl, SelectServiceUrl, SelectSoapActionName, UpdateServiceUrl, UpdateSoapActionName)
    
    function Add-Urn
    {
        param([string] $urn)

        if(-not $Global:urns.Contains($urn))
        {
            $Global:urns += $urn
        }
    }


    function Lookup-LocationInfo
    {
        param($urn)

        $formCabPaths = @($Global:cabPaths | ?{$_.ColumnA -eq $urn})
    
        $tempLocations = @()
        $tempSiteUrls = @()
    
        $retVal = New-Object PSObject -Property @{ 
                                                    XSNUrls = [string]::Empty
                                                    XSNUrlCount = 0
                                                    SiteUrls = [string]::Empty
                                                 }

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
            $tempLocations += [string]::Format("{0}/{1}", $file.DirName, $file.LeafName)
            
            if($tempSiteUrls -notcontains $file.SiteUrl)
            {
                $tempSiteUrls += $file.SiteUrl
            }
        }

        $retVal.XSNUrls = [string]::Join(";", $tempLocations)
        $retVal.XSNUrlCount = $tempLocations.Count
        $retVal.SiteUrls = [string]::Join(";", $tempSiteUrls)
       
        return $retVal
    }

    function Lookup-XSNSiteIds
    {
        param($urn)

        $formCabPaths = @($Global:cabPaths | ?{$_.ColumnA -eq $urn})
    
        $tempSiteIds = @()

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
            $tempSiteIds += $file.SiteId
        }
       
        return $tempSiteIds
    }


    function Get-SoapConnections
    {
        param([PSObject]$currentObject) # the XSN form template to process

        $urls = @()

        foreach($item in $Global:soapConnections)
        {
            if($item.ColumnA -ne $currentObject.URN)
            {
                continue
            }

            if(-not $supportedWebServices.ContainsKey($item.ColumnD))
            {
                $urls += [string]::Format("{0}?{1}", $item.ColumnC, $item.ColumnD)
            }
        }

        foreach($item in $Global:restConnections)
        {
            if($item.ColumnA -ne $currentObject.URN)
            {
                continue
            }

            if ($item.ColumnC -eq "") { $item.ColumnC = "connection in form" }
            
            $urls += [string]::Format("REST:{0}", $item.ColumnC)
        }

        foreach($item in $Global:udcConnections)
        {
            if($item.ColumnA -ne $currentObject.URN)
            {
                continue
            }
            
            if(-not $supportedWebServices.ContainsKey($item.ColumnD))
            {
                $siteIds = @(Lookup-XSNSiteIds $currentObject.URN)             
                $udcxFilePath = [string]::Format("*{0}", $item.ColumnC)                

                #Query the list of UDCX Connection File records for the ones expected by this deployed form instance
                # Notes: 
                #  - XSN Forms can be deployed to a different SPWeb than the one that contains the UDCX file; however, both SPWebs must be within the same SPSite (so we compare SiteIds)
                #  - It is entirely possible that the expected UDCX file is not present (e.g., not deployed, moved, renamed, deleted, etc.)
                $udcxCalls = $udcxRecords | ? { $siteIds -contains $_.SiteId -and $_.RelativeUrl -like  $udcxFilePath}

                foreach($udcxCall in $udcxCalls)
                {
                    if(-not $supportedWebServices.ContainsKey($udcxCall.SelectSoapActionName))
                    {
                        $urls += [string]::Format("{0}?{1}", $udcxCall.SelectServiceUrl, $udcxCall.SelectSoapActionName)
                    }
                }                       
            }
        }
        
        $currentObject.UnsupportedSoapCallsCount = $urls.Count
        $currentObject.UnsupportedSoapCalls = [string]::Join(";", $urls) -replace '[\r\n]',''
    }    

    function Get-DataConnections
    {
        param([PSObject]$currentObject)

        $urls = @()

        foreach($item in $Global:dataConnections)
        {
            if($item.ColumnA -ne $currentObject.URN)
            {
                continue
            }

            if (-not $urls.Contains([string]::Format("{0}", $item.ColumnC)))
            { 
              $urls += [string]::Format("{0}", $item.ColumnC)
            }
        }

        foreach($item in $Global:adoConnections)
        {
            if($item.ColumnA -ne $currentObject.URN)
            {
                continue
            }

            if (-not $urls.Contains([string]::Format("{0}", $item.ColumnC)))
            { 
              $urls += [string]::Format("{0}", $item.ColumnC)
            }
        }
    
        $currentObject.UnsupportedDataConnectionInstances = $urls.Count
        $currentObject.UnsupportedDataConnectionTypes = [string]::Join(";", $urls) -replace '[\r\n]',''
    }

    function Get-ManagedCode
    {
        param([PSObject]$currentObject)

        foreach($item in $Global:managedCode)
        {
            if($item.ColumnA -ne $currentObject.URN)
            {
                continue
            }
    
            $currentObject.ManagedCode = $true
        
            if([int]$item.ColumnD -eq 2)
            {
                $currentObject.ManagedCodeState = "RemediationRequired"
            }
            else
            {
                $currentObject.ManagedCodeState = "ValidationRequired"
            }    

            break
        }
    }

    function Populate-Collections
    {
        foreach($row in Import-CSV $Global:reportFolder\InfoPathFileInfo.log)
        {
            if(-not $fileInfo.ContainsKey($row.Id))
            {
                $fileInfo.Add($row.Id, $row)    
            }
        }

        foreach($item in $infoPathRecords)
        {
            switch($item.ColumnB)
            {
                "PublishUrl" { $Global:publishUrls += $item }

                "CabPath" { $Global:cabPaths += $item }

                "Mode" { $Global:modes += $item }

                "ProductVersion" { $Global:productVersions += $item }

                "SoapConnection" { 
                                    if($item.ColumnD -ne "?")
                                    { 
                                        $Global:soapConnections += $item
                                        Add-Urn($item.ColumnA)
                                    } 
                                 }

                "RESTConnection" {                                      
                                    $Global:restConnections += $item
                                    Add-Urn($item.ColumnA)                                    
                                 }

                "UdcConnection" {
                                    $Global:udcConnections += $item
                                    Add-Urn($item.ColumnA)
                                 }

                "AdoConnection" {
                                    $Global:adoConnections += $item
                                    Add-Urn($item.ColumnA)
                                 }
                                 
                "ManagedCode"    {
                                    $Global:managedCode += $item 
                                    Add-Urn($item.ColumnA)
                                 }

                "DataConnection" {
                                    $Global:dataConnections += $item 
                                    Add-Urn($item.ColumnA)
                                 }
            }
        }
    }


    function Get-FormLocations
    {
        param([PSObject]$currentObject)

        $locationInfo = Lookup-LocationInfo $currentObject.URN

        $currentObject.URLs = $locationInfo.XSNUrls -replace '[\r\n]',''
        $currentObject.InstanceCount = $locationInfo.XSNUrlCount
        $currentObject.SiteUrls = $locationInfo.SiteUrls
    }


    function Get-Mode
    {
        param([PSObject]$currentObject)

        foreach($mode in $Global:modes)
        {
            if($mode.ColumnA -ne $currentObject.URN)
            {
                continue
            }    

            $currentObject.Mode = $mode.ColumnC

            break
        }
    }


    function Get-ProductVersion
    {
        param([PSObject]$currentObject)

        foreach($version in $Global:productVersions)
        {
            if($version.ColumnA -ne $currentObject.URN)
            {
                continue
            }    

            $currentObject.ProductVersion = $version.ColumnC

            break
        }
    }


    function Parse-Report
    {
        Populate-Collections    

        foreach($urn in $Global:urns)
        {
            if($reportOutput.ContainsKey($urn))
            {
                continue
            }

            $currentObject = New-Object PSObject -Property @{ 
                                                               URN = $urn
                                                               InstanceCount = 0
                                                               SiteUrls = [string]::Empty
                                                               URLs = [string]::Empty
                                                               UnsupportedSoapCalls = [string]::Empty
                                                               UnsupportedSoapCallsCount = 0                                                              
                                                               UnsupportedDataConnectionTypes = [string]::Empty
                                                               UnsupportedDataConnectionInstances = 0
                                                               ManagedCode = $false
                                                               ManagedCodeState = [string]::Empty
                                                               Mode = [string]::Empty
                                                               ProductVersion = [string]::Empty
                                                             }

            Get-SoapConnections ([REF]$currentObject)
            Get-ManagedCode ([REF]$currentObject)
            Get-DataConnections ([REF]$currentObject)

            $isSoapUnsupported = $currentObject.UnsupportedSoapCallsCount -gt 0
            $isDataConUnsupported = $currentObject.UnsupportedDataConnectionInstances -gt 0
            $isManaged = $currentObject.ManagedCode -eq $true
            
            if($isSoapUnsupported -or $isManaged -or $isDataConUnsupported)
            {
                Get-FormLocations ([REF]$currentObject)
                Get-Mode ([REF]$currentObject)
                Get-ProductVersion ([REF]$currentObject)

                if(-not $reportOutput.ContainsKey($urn))
                {
                    $reportOutput.Add($urn, $currentObject);
                }
            }        
        }

        if($reportOutput.Count -gt 0)
        {
            $reportOutput.Values | 
                            ?{$_.Mode.ToLower() -ne "client"} | 
                            Select URN, InstanceCount, SiteUrls, URLs, UnsupportedSoapCalls, UnsupportedSoapCallsCount, UnsupportedDataConnectionTypes, UnsupportedDataConnectionInstances, ManagedCode, ManagedCodeState, Mode, SolutionFormatVersion, ProductVersion | 
                            Export-CSV "$Global:exportFolder\$outputFile" -NoTypeInformation 
        }
        else
        {
             "No impacted InfoPath forms were located." | Out-File "$Global:exportFolder\$outputFile"   
        }
    }

    Parse-Report

    $retVal = $true
}
catch [Exception]
{
    $retVal = $_.CategoryInfo.Activity + " : " + $_.Exception.Message + $_.InvocationInfo.PositionMessage
    $retVal = $retVal.Replace("`"", "`'")
    $retVal = $retVal.Replace("'", "``'")
}

return $retVal