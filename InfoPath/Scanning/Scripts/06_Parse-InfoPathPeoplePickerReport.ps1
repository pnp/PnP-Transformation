<#
.SYNOPSIS
Parses the output from 02_Scrape-InfoPathFiles.ps1 to generate a customer friendly people picker instances report.

.DESCRIPTION
Execute the script to parse out the report.csv file

.Example
.\InfoPath_Scanner.ps1

#>
try
{
    $retVal = $false

    $ErrorActionPreference = "STOP"

    $Global:exportFolder = "c:\Scanner\InfoPath"

    $Global:reportFolder = "$Global:exportFolder\xsn"
    $Global:udcxFolder = "$Global:exportFolder\udcx"
    $Global:cabPaths = @()
    $Global:modes = @()
    $Global:peoplePickers = @()
    $Global:urns = @()

    $reportOutput = @{}
    $fileInfo = @{}

    $fileDateTime = [DateTime]::Now.ToString("MMddyyyy_hhmmss")  #Ensure the Output and Error Log Files have the same Date/Time name
    $outputFile = [string]::Format("InfoPathForms_PeoplePicker_{0}.csv", $fileDateTime)


    if (-not (Test-Path $Global:exportFolder)) 
    { 
        New-Item $Global:exportFolder -ItemType Directory | Out-Null
    }
    
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
    
        $returnVal = New-Object PSObject -Property @{ 
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

        $returnVal.XSNUrls = [string]::Join(";", $tempLocations)
        $returnVal.XSNUrlCount = $tempLocations.Count
        $returnVal.SiteUrls = [string]::Join(";", $tempSiteUrls)
       
        return $returnVal
    }

    
    function Get-PeoplePickers
    {
        param([PSObject]$currentObject)

        $count = 0

        foreach($item in $Global:peoplePickers)
        {
            if($item.ColumnA -ne $currentObject.URN)
            {
                continue
            }

            $count++
        }  
        
        $currentObject.PeoplePickersCount = $count
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

        $infoPathRecords = Import-CSV $Global:reportFolder\InfoPathScraper_report.csv -Header ColumnA, ColumnB, ColumnC, ColumnD
        foreach($item in $infoPathRecords)
        {
            if ($item.ColumnC -eq "{{61e40d31-993d-4777-8fa0-19ca59b6d0bb}}")
            {
                $Global:peoplePickers += $item
                Add-Urn($item.ColumnA)
            }

            if ($item.ColumnB -eq "CabPath")
            {
                $Global:cabPaths += $item
            }
        }
        $infoPathRecords = $null
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
                                                               PeoplePickersCount = 0
                                                               Mode = [string]::Empty
                                                             }


            Get-PeoplePickers ([REF]$currentObject)
            if ($currentObject.PeoplePickersCount -gt 0)
            {
                Get-FormLocations ([REF]$currentObject)
                Get-Mode ([REF]$currentObject)
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
                            Select URN, InstanceCount, SiteUrls, URLs, PeoplePickersCount | 
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