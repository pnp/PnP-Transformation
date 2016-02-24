<#
.SYNOPSIS
Parses information out of all the XSN/UDCX files in the farm and generates output files.

.DESCRIPTION
Execute the script to scan all content databases and feature folders for impacted InfoPath templates.

Typical errors one might see with this script are:
- "Microsoft.SharePoint.SPException: Access to this Web site has been blocked. Please contact the administrator to resolve this problem."
- "System.Management.Automation.RuntimeException: You cannot call a method on a null-valued expression."
- "Cannot find an SPSite object with Id or Url: <site id>. ---> System.IO.FileNotFoundException: The site with the id <site id> could not be found"

These errors all indicate that the site where the XSN/UDCX file is residing either is deleted (bottom 2 errors) or is locked (first error).


.Example
.\01_Get-InfoPathFiles.ps1

#>

$getInfoPathFiles = {
    try
    {
        $ErrorActionPreference = "STOP"
        $retVal = $false

        $localPath = "c:\Scanner\InfoPath"   

        $xsnFileExtension = "xsn"
        $udcxFileExtension = "udcx"

        $fileDateTime = [DateTime]::Now.ToString("MMddyyyy_hhmmss")  #Ensure the Output and Error Log Files have the same Date/Time name
        $errorFile = [string]::Format("{0}\{1}\InfoPath_Errors_{2}.csv", $localPath, $xsnFileExtension, $fileDateTime)

        #$filePath = $localPath + "\files.txt"
        #$reportPath = $localPath + "\report.csv"
        #$errorLogPath =  $localPath + "\error.log"
        
        $infoPathFileInfo = $localPath + "\" + $xsnFileExtension + "\InfoPathFileInfo.log"
        $udcxFileInfo = $localPath + "\" + $udcxFileExtension + "\InfoPathFileInfo.log"

    
        if((Get-PSSnapin | ?{$_.Name -match "Microsoft.SharePoint.PowerShell"}) -eq $null)
        {
            Add-PSSnapin Microsoft.SharePoint.PowerShell
        }   
    
        ## the query to be execute against each content database
        $query = "select SiteId, WebId, DirName, LeafName, Id from AllDocs (nolock)
        where Extension = '{0}'
        and DeleteTransactionId = 0x
        and HasStream = 1
        order by SiteId, WebId, DirName"                        

        ## Function
        ## execute the predefined query for a given databases
        function getSiteRecords($database, $sqlQuery)
        {
            $ds = $null

            try
            {
            
                $contentdata = $null
                $conn = new-object System.Data.SqlClient.SqlConnection
                $conn.ConnectionString = $database.DatabaseConnectionString
                $conn.Open() | out-null
                $sqlCmd = New-Object System.Data.SqlClient.SqlCommand
                $sqlCmd.CommandText = $sqlQuery
                #$sqlCmd.CommandTimeout = 180
                $sqlCmd.Connection = $conn
                $contentdata = New-Object System.Data.SqlClient.SqlDataAdapter
                $ds = New-Object System.Data.DataSet
                $contentdata.SelectCommand = $sqlCmd
                $contentdata.Fill($ds) | out-null
                $conn.Close() | out-null            
            }
            catch [Exception]
            {
                Write-Host "Error" $_.Exception.ToString()
                if ($conn -ne $null) {$conn.Close()}
            }
              
            $ds
        }

        function Error-To-File($obj, $output)
        {
           $output | Add-Member -MemberType NoteProperty -Name Error -Value $obj.Exception.ToString()
           $output | Export-Csv -notypeinformation -Path $errorFile  -Append
        }

        function Download-Files($fileNameExtension)
        {        
            $rowObjects = @()
        
            ###############################################################################################################
            ## Scan all content databases for XSN templates/UDCX Files
            ###############################################################################################################
            ## iterate all content databases
            foreach($db in (Get-SPContentDatabase))
            {
                ## query the database
                $sqlQuery = [string]::Format($query, $fileNameExtension)
                $data = getSiteRecords $db $sqlQuery

                if ($data -ne $null -and $data.Tables.Count -gt 0)
                {
                    ## convert data to PowerShell object
                    foreach($row in $data.Tables[0].Rows)
                    {
                        $t++
                        $rowObject = New-Object -TypeName PSObject
                        $rowObject | Add-Member -MemberType NoteProperty -Name SiteId -Value $row[0]
                        $rowObject | Add-Member -MemberType NoteProperty -Name WebId -Value $row[1]
                        $rowObject | Add-Member -MemberType NoteProperty -Name DirName -Value $row[2]
                        $rowObject | Add-Member -MemberType NoteProperty -Name LeafName -Value $row[3]
                        $rowObject | Add-Member -MemberType NoteProperty -Name Id -Value $row[4]
                        $rowObjects += $rowObject
                    }
                }
            }   
   
            ###############################################################################################################
            ## Download and scan every XSN template/UDCX file found in a content DB
            ###############################################################################################################
            ## now we've got a list of the xsn files to download....
            ## start processing each file
            $web = $null
            $site = $null
            $lastWebId = ""
            $siteUrl = ""
            $webUrl = ""
            $localFilePath = $localPath + "\" + $fileNameExtension

            foreach ($doc in $rowObjects)
            {     
                try
                {   
                    ## Has the web changed?
                    if ($doc.WebId -ne $lastWebId)
                    {
                        ## web has changed - load it
                        ## dispose of the last web and site if needed
                        if ($web -ne $null) {$web.Dispose()}
                        if ($site -ne $null) {$site.Dispose()}

                        ## get the site
                        $site = Get-SPSite -Identity $doc.SiteId
                        $siteUrl = $site.Url
        
                        ## get the web
                        $web = $site.AllWebs | ? {$_.id -eq $doc.WebId}
                        $webUrl = $web.Url
                        $lastwebId = $web.id            
                    }
                    
                    $doc | Add-Member -MemberType NoteProperty -Name SiteUrl -Value $siteUrl
                    $doc | Add-Member -MemberType NoteProperty -Name WebUrl -Value $webUrl
    
                    ## we've got the web so get the doc
                    $file = $null
                    $file = $web.GetFile($doc.Id)
                    if ($file -ne $null -and $file.Exists)
                    {
                        try
                        {
                            $bytes = $file.OpenBinary();             
                  
                            # Download the file to the path
                            $docId = $doc.Id
                            $localfile = $localFilePath + "\" + $docId + "\" + $doc.LeafName
                            New-Item "$localFilePath\$docId" -type directory -force | Out-Null 
                            [System.IO.FileStream] $fs = new-object System.IO.FileStream($localfile, "OpenOrCreate") 
                            $fs.Write($bytes, 0 , $bytes.Length) 
                            $fs.Close()                               
                        }
                        catch [Exception]
                        {
                            Error-To-File $_ $doc
                        }
                    }
                }
                catch [Exception]
                {
                    Error-To-File $_ $doc
                }               
                
            } #foreach ($doc in $rowObjects)

            if ($web -ne $null) {$web.Dispose()}
            if ($site -ne $null) {$site.Dispose()}
            $web = $null
            $site = $null
            
            $rowObjects  
        } # downloadFiles  z

        
        $xsnObjects = Download-Files $xsnFileExtension
        $xsnObjects | Export-CSV -notypeinformation $infoPathFileInfo
        $xsnObjects = $null

        $udcxObjects = Download-Files $udcxFileExtension
        $udcxObjects | Export-CSV -notypeinformation $udcxFileInfo
        $udcxObjects = $null

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

Powershell.exe -command $getInfoPathFiles
