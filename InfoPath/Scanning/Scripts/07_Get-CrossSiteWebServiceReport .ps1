<#
.SYNOPSIS
Parses the output from 02_Scrape-InfoPathFiles.ps1 to generate a customer friendly report.

The report identifies those InfoPath form (XSN) instances that use a UDCX file to call an 
OOB Web Service that is supported in SharePoint Online (MT) and (vNext).  It examines each call
and flags those that are considered to be cross-site calls. Cross-site OOB Web Service calls 
will fail in SPO-MT/vNext; as such, the offending forms/udcx need to be remediated.

.DESCRIPTION
Execute the script to generate a report of Cross-Site (OOB) Web Service Calls

.Example
.\07_Get-CrossSiteWebServiceReport.ps1

#>
try
{
    #======================================================================================
    #ACTION ITEMS: Please initialize these settings with values that suit your environment

    #specfiy the location of your InfoPath Report folder
    $Global:exportFolder = "c:\Scanner\InfoPath"

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
    $Global:urns = @()

    #Each member of this dictionary represents a row in the output report
    $reportOutput = @{}
    #each member of this dictionary represents a row from the InfoPathFileInfo.log
    $fileInfo = @{}

    #each member of this dictionary represents an OOB Web Services that is supported in SPO-MT (and SPO-vNext)
    $supportedWebServices = @{
                                "GetUserProfileByName" = "userprofileservice.asmx"
                                "GetGroupCollectionFromWeb" = "usergroup.asmx"	
                                "GetUserCollectionFromGroup" = "usergroup.asmx"
                                "GetCommonManager" = "userprofileservice.asmx"
                                "GetCommonMemberships" = "userprofileservice.asmx"
                                "GetUserMemberships" = "userprofileservice.asmx"
                                "GetUserPropertyByAccountName" = "userprofileservice.asmx"
                                "CheckInFile" = "lists.asmx"
                                "CheckOutFile" = "lists.asmx"
                                "GetUserCollectionFromSite" = "usergroup.asmx"
                            }

    $fileDateTime = [DateTime]::Now.ToString("MMddyyyy_hhmmss")  #Ensure the Output and Error Log Files have the same Date/Time name

    $outputFile = [string]::Format("InfoPathForms_CrossSiteWebService_{0}.csv", $fileDateTime)
    $errorFile = [string]::Format("InfoPathForms_CrossSiteWebService_Errors_{0}.csv", $fileDateTime)


    if (-not (Test-Path $Global:exportFolder)) 
    { 
        New-Item $Global:exportFolder -ItemType Directory | Out-Null
    }


    #This is the Form Template (XSN) contents/attributes report
    $infoPathRecords = Import-CSV $Global:reportFolder\InfoPathScraper_report.csv -Header ColumnA, ColumnB, ColumnC, ColumnD
    
    #This is the UDCX File instance report
    $udcxRecords = @(Import-CSV $Global:udcxFolder\UDCXReport.csv | Select WebId, SiteUrl, RelativeUrl, SelectServiceUrl, SelectSoapActionName)
    

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

        foreach($mode in $Global:modes)
        {
            if($mode.ColumnA -ne $currentObject.XsnUrn)
            {
                continue
            }    

            $currentObject.XsnMode = $mode.ColumnC

            break
        }
    }

    function Lookup-XsnLocationInfo
    {
        param($urn)

        $formCabPaths = @($Global:cabPaths | ?{$_.ColumnA -eq $urn})
    
        $tempLocations = @()
        $tempSiteUrls = @()
    
        $retVal = New-Object PSObject -Property @{ 
                                                    XsnFileUrls = [string]::Empty
                                                    XsnUrlCount = 0
                                                    XsnSiteUrls = [string]::Empty
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
            
            $tempSiteUrls += $file.SiteUrl
        }

        $retVal.XsnFileUrls = [string]::Join(";", $tempLocations)
        $retVal.XsnUrlCount = $tempLocations.Count
        $retVal.XsnSiteUrls = [string]::Join(";", $tempSiteUrls)
       
        return $retVal
    }

    function Lookup-XsnWebIds
    {
        param($urn)

        $formCabPaths = @($Global:cabPaths | ?{$_.ColumnA -eq $urn})
    
        $tempWebIDs = @()

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
            $tempWebIDs += $file.WebId
        }
       
        return $tempWebIDs
    }

    #Parses a UDCX ServiceUrl and extracts the Url of the hosting SPSite
    function Get-ServiceSiteUrl
    {
        param([string] $url)

        $pos = $url.IndexOf("/_vti_bin")
        if ($pos -lt 0)
        {
            # this Url does not contain a valid Service call
            return ""
        }

        # We have a valid service Url; this can be either an SPSite (https://portal.constoso.com/sites/siteX) or an SPWeb (https://portal.constoso.com/sites/siteX/webX)
        # we assume it can contain zero or one managedPath
        # We want to return only the SPSite Url...

        # Check for each of the supported managed paths 
        foreach ($mp in $Global:managedPaths)
        {
            if ($mp.StartsWith("/") -eq $false)
            {
                $mp = "/" + $mp
            }
            if ($mp.EndsWith("/") -eq $false)
            {
                $mp = $mp + "/"
            }

            $pos = $url.IndexOf($mp)
            if ($pos -gt -1)
            {
                $webApp = $url.Substring(0, $pos + $mp.Length)
                $remainder = $url.Substring($pos + $mp.Length)

                $pos2 = $remainder.IndexOf("/")
                $siteName = $remainder.Substring(0, $pos2)

                $siteUrl = $webApp + $siteName
                return $siteUrl
            }
        }

        # at this point, we must be dealing with the root site collection of a web app (https://portal.constoso.com, (https://my.constoso.com, etc.)
        # 
        # In addition, we might reach this point for managed paths that have not been entered into $Global:managedPaths 
        $pos = $url.IndexOf("/_vti_bin")
        return $url.Substring(0, $pos)
    }

    function Get-FormLocations
    {
        param([PSObject]$currentObject)

        $locationInfo = Lookup-XsnLocationInfo $currentObject.XsnUrn

        $currentObject.XsnInstanceCount = $locationInfo.XsnUrlCount
        $currentObject.XsnFileUrls = $locationInfo.XsnFileUrls.ToLower() -replace '[\r\n]',''
        $currentObject.XsnSiteUrls = $locationInfo.XsnSiteUrls.ToLower()
    }

    function Get-UdcxConnections
    {
        param([PSObject]$currentObject)

        $udcxSiteUrls = @()    
        $udcxFileUrls = @()    
        $udcxSupportedServiceSiteUrls = @()
        $udcxSupportedServiceUrls = @()

        foreach($item in $Global:udcxConnections)
        {
            if($item.ColumnA -ne $currentObject.XsnUrn)
            {
                continue
            }
            
            if($supportedWebServices.ContainsKey($item.ColumnD))
            {
                #the current web service is supported in vNext...
                $webIds = @(Lookup-XsnWebIds $currentObject.XsnUrn)             
                $udcxFilePath = [string]::Format("*{0}", $item.ColumnC)                
                $udcxCalls = $udcxRecords | ? { $webIds -contains $_.WebId -and $_.RelativeUrl -like  $udcxFilePath}

                foreach($udcxCall in $udcxCalls)
                {
                    if($supportedWebServices.ContainsKey($udcxCall.SelectSoapActionName))
                    {
                        $udcxSiteUrls += [string]::Format("{0}", $udcxCall.SiteUrl.ToLower())
                        $udcxFileUrls += [string]::Format("{0}", $udcxCall.RelativeUrl.ToLower())

                        $udcxSupportedServiceSiteUrls += Get-ServiceSiteUrl($udcxCall.SelectServiceUrl.ToLower())
                        $udcxSupportedServiceUrls += [string]::Format("{0}?{1}", $udcxCall.SelectServiceUrl.ToLower(), $udcxCall.SelectSoapActionName.ToLower())
                    }
                }
            }
        }

        # Assert: 
        #$udcxSupportedServiceSiteUrls.Count == $udcxFileUrls.Count == $udcxSupportedServiceUrls.Count
        
        $currentObject.UdcxFileUrlsCount = $udcxFileUrls.Count
        if ($udcxFileUrls.Count -gt 0)
        {
            $currentObject.UdcxSiteUrls = [string]::Join(";", ($udcxSiteUrls)).ToLower() -replace '[\r\n]',''
            $currentObject.UdcxFileUrls = [string]::Join(";", ($udcxFileUrls)).ToLower() -replace '[\r\n]',''
        }
        
        $currentObject.UdcxSupportedServiceUrlsCount = $udcxSupportedServiceUrls.Count
        if ($udcxSupportedServiceUrls.Count -gt 0)
        {
            $currentObject.UdcxSupportedServiceSiteUrls = [string]::Join(";", ($udcxSupportedServiceSiteUrls)).ToLower() -replace '[\r\n]',''
            $currentObject.UdcxSupportedServiceUrls = [string]::Join(";", ($udcxSupportedServiceUrls)).ToLower() -replace '[\r\n]',''
        }
    }    

    function Get-Remediation
    {
        param([PSObject]$currentObject)

        [string] $xsnSiteUrl = $currentObject.XsnSiteUrls
        [string] $udcxSiteUrl = $currentObject.UdcxSiteUrls
        [string] $serviceSiteUrl = $currentObject.UdcxSupportedServiceSiteUrls

        # If the SPSite of the XSN equals the SPSite of the serviceCall, 
        # then a cross-site call is not taking place and no remediation is needed.
        if ($xsnSiteUrl.Trim().Equals($serviceSiteUrl.Trim()))
        {
            $currentObject.Remediation = "None"
            return
        }

        # at this point, we have a cross-site service call...

        # If the SPSite of the XSN (i.e., form) equals the SPSite of the udcx file, 
        # the remediation is to simply edit the serviceUrl in the udcx file to use the SPSite of the XSN (i.e., form) 
        if ($xsnSiteUrl.Trim().Equals($udcxSiteUrl.Trim()))
        {
            $currentObject.Remediation = "Update UDCX"
            return
        }
        # If the SPSite of the XSN does not equal the SPSite of the udcx file, 
        # the remediation is to republish the XSN (i.e., form) instance.
        else
        {
            $currentObject.Remediation = "Republish Form"
            return
        }
    }

    function Process-Collections
    {
        foreach($urn in $Global:urns)
        {

            # skip this form if we have already added it to the report
            if($reportOutput.ContainsKey($urn))
            {
                continue
            }

            $currentObject = New-Object PSObject -Property @{ 
                XsnUrn = $urn
                XsnMode = [string]::Empty
                XsnInstanceCount = 0
                XsnSiteUrls = [string]::Empty
                XsnFileUrls = [string]::Empty
                UdcxFileUrlsCount = 0
                UdcxSiteUrls = [string]::Empty
                UdcxFileUrls = [string]::Empty
                UdcxSupportedServiceUrlsCount = 0                                                              
                UdcxSupportedServiceSiteUrls = [string]::Empty
                UdcxSupportedServiceUrls = [string]::Empty
                UdcxNeedsToBeVerified = [string]::Empty
                Remediation = [string]::Empty
                }

            Get-UdcxConnections ([REF]$currentObject)
            
            $hasUdcxConnection = $currentObject.UdcxFileUrlsCount -gt 0
            $hasSupportedServiceCall = $currentObject.UdcxSupportedServiceUrlsCount -gt 0
            
            if($hasUdcxConnection -or $hasSupportedServiceCall)
            {
                Get-FormLocations ([REF]$currentObject)
                Get-XsnMode ([REF]$currentObject)

                #At this point, we output a row for each XSN Instance, as well as for each Service call
                # We effectively have two sets of parallel collections:
                # - XsnSiteUrls[n] and XsnFileUrls[n], where 'n' = XsnInstanceCount
                # - UdcxFileUrls[x], UdcxSupportedServiceSiteUrls[x], and UdcxSupportedServiceUrls[x], where 'x' = UdcxFileUrlsCount (assert UdcxSupportedServiceUrlsCount == UdcxFileUrlsCount)
                #
                $local:xsnSiteUrls = $currentObject.XsnSiteUrls -split ";"
                $local:xsnFileUrls = $currentObject.XsnFileUrls -split ";"
                $local:udcxSiteUrls = $currentObject.UdcxSiteUrls -split ";"
                $local:udcxFileUrls = $currentObject.UdcxFileUrls -split ";"
                $local:udcxSupportedServiceSiteUrls = $currentObject.UdcxSupportedServiceSiteUrls -split ";"
                $local:udcxSupportedServiceUrls = $currentObject.UdcxSupportedServiceUrls -split ";"

                #expand/denormalize the parallel collections of each row
                for ($i=0; $i -lt $currentObject.XsnInstanceCount; $i++)
                {
                    for ($j=0; $j -lt $currentObject.UdcxFileUrlsCount; $j++)
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
                            UdcxSupportedServiceUrlsCount = $currentObject.UdcxSupportedServiceUrlsCount                                                       
                            UdcxSupportedServiceSiteUrls = [string]::Empty
                            UdcxSupportedServiceUrls = [string]::Empty
                            UdcxNeedsToBeVerified = $currentObject.UdcxNeedsToBeVerified
                            Remediation = [string]::Empty
                            }

                        $tempObject.XsnSiteUrls = $local:xsnSiteUrls[$i]
                        $tempObject.XsnFileUrls = $local:xsnFileUrls[$i]

                        $tempObject.UdcxSiteUrls = $local:udcxSiteUrls[$j]
                        $tempObject.UdcxFileUrls = $local:udcxFileUrls[$j]

                        $tempObject.UdcxSupportedServiceSiteUrls = $local:udcxSupportedServiceSiteUrls[$j]
                        $tempObject.UdcxSupportedServiceUrls = $local:udcxSupportedServiceUrls[$j]

                        # Now that the row has been normalized, we can compute the remediation for the form/udcx represented by the row
                        Get-Remediation ([REF]$tempObject)

                        #Add the row to the output collection
                        $reportOutput.Add($urn + '#' + $i.ToString() + '.' + $j.ToString(), $tempObject);
                    }
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
                Select -Unique Remediation, XsnUrn, XsnInstanceCount, XsnSiteUrls, XsnFileUrls, UdcxSiteUrls, UdcxFileUrls, UdcxSupportedServiceSiteUrls, UdcxSupportedServiceUrls | 
                Sort-Object -CaseSensitive -Property XsnUrn, XsnSiteUrls, XsnFileUrls, UdcxSiteUrls, UdcxFileUrls, UdcxSupportedServiceSiteUrls, UdcxSupportedServiceUrls |
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