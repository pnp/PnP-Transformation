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

    $fileDateTime = [DateTime]::Now.ToString("MMddyyyy_hhmmss")  #Ensure the Output and Error Log Files have the same Date/Time name
    $errorFile = [string]::Format("{0}\InfoPath_Errors_{1}.csv", $localPath, $fileDateTime)

    $xsnFileExtension = "xsn"
    $xsnFolder = Join-Path $localPath $xsnFileExtension
    $filePath = $xsnFolder + "\InfoPathScraper_files.txt"
    $reportPath = $xsnFolder + "\InfoPathScraper_report.csv"
    $errorLogPath =  $xsnFolder + "\InfoPathScraper_error.log" 
    $xsnInfoPath = $xsnFolder + "\InfoPathFileInfo.log"

    ## Run the scraper utility over everything
    cmd.exe /c "dir /s/b $xsnFolder\*.xsn > $filePath"

    #InfoPathScraper.exe was failing with an error stating $filePath could not be found, I think that is a timing issue on larger 
    #farms where the above dir call hasn't quite closed the handle to $filepath. Putting this hacky workaround in place to
    #try and avoid that problem.
    Start-Sleep -Seconds 60     

    cmd.exe /C "$localPath\InfoPathScraper\infopathscraper.exe /csv /filelist $filePath /outfile $reportPath /append 2> $errorLogPath"
    
    $retVal = $true
}
catch [Exception]
{
    $retVal = $_.CategoryInfo.Activity + " : " + $_.Exception.Message + $_.InvocationInfo.PositionMessage
    $retVal = $retVal.Replace("`"", "`'")
    $retVal = $retVal.Replace("'", "``'")
}

return $retVal