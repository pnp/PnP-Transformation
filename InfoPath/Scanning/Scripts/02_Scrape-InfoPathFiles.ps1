<#
.SYNOPSIS
Analyzes InfoPath files and dumps relevant information.

.DESCRIPTION
Execute the script to 'scrape' all the exported XSN files and dump out the relevent information needed to analyze the forms.

.Example
.\02_Scrape-InfoPathFiles.ps1

#>

try
{
    $ErrorActionPreference = "STOP"
    $retVal = $false

    $localPath = "c:\Scanner\InfoPath"   
    $scraperPath = "c:\Scanner\InfoPath\InfoPathScraper"

    $fileDateTime = [DateTime]::Now.ToString("MMddyyyy_hhmmss")  #Ensure the Output and Error Log Files have the same Date/Time name
    $errorFile = [string]::Format("{0}\InfoPath_Errors_{1}.csv", $localPath, $fileDateTime)

    $xsnFileExtension = "xsn"
    $xsnFolder = Join-Path $localPath $xsnFileExtension
    $filePath = $xsnFolder + "\InfoPathScraper_files.txt"
    $reportPath = $xsnFolder + "\InfoPathScraper_report.csv"
    $errorLogPath =  $xsnFolder + "\InfoPathScraper_error.log" 
    $xsnInfoPath = $xsnFolder + "\InfoPathFileInfo.log"


    $StartDate=(GET-DATE)
        
    #grab the file list
    $xsnFiles = Get-ChildItem $xsnFolder -Recurse -Filter *.xsn
    #run scraper file per file
    foreach($xsn in $xsnFiles)
    {
        [string] (& $scraperPath\infopathscraper.exe /csv /file `"$($xsn.FullName)`" /outfile `"$reportPath`" /append  2>&1) | out-file $errorLogPath -Append
    } 

    $endDate=(GET-DATE)
    NEW-TIMESPAN –Start $StartDate –End $EndDate | out-file $errorLogPath -Append   

    $retVal = $true
}
catch [Exception]
{
    $retVal = $_.CategoryInfo.Activity + " : " + $_.Exception.Message + $_.InvocationInfo.PositionMessage
    $retVal = $retVal.Replace("`"", "`'")
    $retVal = $retVal.Replace("'", "``'")
}

return $retVal