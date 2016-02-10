<#
.SYNOPSIS
Parses the output from 04_Parse-InfoPathReport.ps1 to generate a InfoPath form usage information.

.DESCRIPTION
Execute the script to parse out the report.csv file

.Example
.\05_Get-InfoPathUsageInformation.ps1

#>
Add-PSSnapIn Microsoft.SharePoint.Powershell -EA 0

$baseFolder = "c:\Scanner\InfoPath"
$logFileName = "InfoPathFileInfo.log"

$fileDateTime = [DateTime]::Now.ToString("MMddyyyy_hhmmss")  #Ensure the Output and Error Log Files have the same Date/Time name
$outputFile = [string]::Format("{0}\InfoPathFormsUsage_{1}.csv", $baseFolder, $fileDateTime)

$separator = ";"
$itemPattern = "/Item/"
$formPattern = "/Forms/"
$attachmentPattern = "/Attachments/"
$formServerPattern = "/FormServerTemplates/"
$ctsPattern = "/_cts/"

#Create Table object
$table = New-Object system.Data.DataTable “InfoPath”

#Define Columns
$col3 = New-Object system.Data.DataColumn Urn,([string])
$col4 = New-Object system.Data.DataColumn InstanceCount,([int])
$col5 = New-Object system.Data.DataColumn SiteUrls,([string])
$col6 = New-Object system.Data.DataColumn Urls,([string])
$col7 = New-Object system.Data.DataColumn UnsupportedSoapCalls,([string])
$col8 = New-Object system.Data.DataColumn UnsupportedSoapCallsCount,([int])
$col9 = New-Object system.Data.DataColumn UnsupportedDataConnectionTypes,([string])
$col10 =New-Object system.Data.DataColumn UnsupportedDataConnectionInstances,([int])
$col11 = New-Object system.Data.DataColumn ManagedCode,([string])
$col12 = New-Object system.Data.DataColumn ManagedCodeState,([string])
$col13 = New-Object system.Data.DataColumn Mode,([string])
$col14 = New-Object system.Data.DataColumn ProductVersion,([string])
$col15 = New-Object system.Data.DataColumn ListName, ([string])
$col16 = New-Object system.Data.DataColumn ItemCount, ([string])
$col17 = New-Object system.Data.DataColumn LastModifiedDate,([string])

<#$col17 = New-Object system.Data.DataColumn Author,([string])
$col18 = New-Object system.Data.DataColumn SiteOwners,([string])
$col19 = New-Object system.Data.DataColumn SiteAdmins,([string])#>

#Add the Columns
$table.columns.add($col3)
$table.columns.add($col4)
$table.columns.add($col5)
$table.columns.add($col6)
$table.columns.add($col7)
$table.columns.add($col8)
$table.columns.add($col9)
$table.columns.add($col10)
$table.columns.add($col11)
$table.columns.add($col12)
$table.columns.add($col13)
$table.columns.add($col14)
$table.columns.add($col15)
$table.columns.add($col16)
$table.columns.add($col17)


function Download-InfoPathForm ($formURL, $site, $folder, $docID)
{
    $xsnDocIDFolder = [string]::Format("{0}\{1}", $folder, $docID)
    if (-not (Test-Path $xsnDocIDFolder)) 
    { 
        New-Item $xsnDocIDFolder -ItemType Directory | Out-Null
    }

    $web = $site.OpenWeb()
    $file = $web.GetFile($formURL)
    $xsnFilePath = [string]::Format("{0}\{1}", $xsnDocIDFolder, $file.Name)

    $bytes = $file.OpenBinary();
 
    [System.IO.FileStream] $fs = new-object System.IO.FileStream($xsnFilePath, "OpenOrCreate") 
    $fs.Write($bytes, 0 , $bytes.Length) 
    $fs.Close()  
}

function Get-MetaData($URL, $site)
{
    $listName = "N/A"
    $author = "N/A"
    $itemCount = "N/A"
    $lastModifiedDate = "N/A"
    $owners = "N/A"
    $admins = "N/A"
    $ctypeData = @()

    $web = $site.OpenWeb()
    
    if ($URL.Contains($itemPattern) -or $URL.Contains($formPattern) -or $URL.Contains($attachmentPattern))
    {
        $pattern = $itemPattern
        if ($URL.Contains($formPattern))
        {
            $pattern = $formPattern
        }
        elseif ($URL.Contains($attachmentPattern))
        {
            $pattern = $attachmentPattern
        }
        
        $listURL = $URL.Substring(0, $URL.LastIndexOf($pattern))

        try
        {
            $lst = $web.GetList($listURL) | select Author, ItemCount, LastItemModifiedDate
            
            $author = $lst.Author.UserLogin
            $lstName = $lst.Title
            $itemCount = $lst.ItemCount
            $lastModifiedDate = $lst.LastItemModifiedDate
        }
        Catch [system.exception]
        {
            $itemCount = "Error"
        }
    }
    elseif ($URL.Contains($formServerPattern) -or $URL.Contains($ctsPattern))
    {
        $URL = $URL.Substring($URL.LastIndexOf("/")+1)
        $URL = $URL.Replace(".xsn", "")
        $ctLists = $web.Lists.ContentTypes | ? { $_.name -eq $URL } | select parentlist

        foreach ($ctLst in $ctLists)
        {
            $lst = $web.lists | ? { $_.Title -eq $ctLst.ParentList } | select Author, ItemCount, LastItemModifiedDate
            $listName = $ctLst.ParentList 

            if ($author -eq "N/A")
            {
                $author = $lst.Author.UserLogin
                $itemCount = $lst.ItemCount
                $lastModifiedDate = $lst.LastItemModifiedDate
            }
            else
            {
                $ctypeData += @{"ListName"=$listName; "Author"=$lst.Author.UserLogin; "ItemCount"= $lst.ItemCount; "ModifiedDate"= $lst.LastItemModifiedDate}
            }  

        }
    }
    
    # Get site owners
    <#$users = $web.AssociatedOwnerGroup.Users
    foreach($user in $users)
    {
        if ($owners -eq "N/A")
        {
            $owners = $user
        }
        else
        {
            $owners = [string]::Format("{0};{1}", $owners, $user)
        }    

    }

    # Get Site Collection Administrators
    $adminstrators = $web.SiteAdministrators
    foreach($admin in $adminstrators)
    {
        if ($admins -eq "N/A")
        {
            $admins = $admin
        }
        else
        {
            $admins = [string]::Format("{0};{1}", $admins, $admin)
        }    

    }#>

    #return $author,$itemCount,$lastModifiedDate,$owners,$admins, $listName, $ctypeData
    return  $itemCount,$lastModifiedDate,$listName, $ctypeData
}

function Get-Data
{
    $folder = $baseFolder
    $usageInfoPathFolder = [string]::Format("{0}\{1}", $folder, "UsageInfoPaths")
    $logFilePath = [string]::Format("{0}\{1}\{2}", $folder, "xsn", $logFileName)
    $ctypeLists = @()

    if (-not (Test-Path $usageInfoPathFolder)) 
    { 
        New-Item $usageInfoPathFolder -ItemType Directory | Out-Null
    }

    $file = (Get-ChildItem -LiteralPath $folder -Filter "infopathforms_*").Name
    #Read all data 
    $rows = @(import-CSV -LiteralPath([string]::Format("{0}\{1}", $folder, $file)))
    # Read Log Info
    $infoPathFileInfoData = Import-Csv $logFilePath
    
    foreach($r in $rows)
    {
        $isXSNFileDownloaded = $false

        $URLs = $r.Urls -split $separator            

        foreach($URL in $URLs)
        {
            $row = $table.NewRow()

            $row.Urn = $r.Urn
            $row.InstanceCount = 1 
            $row.Urls = $URL
            $row.SiteUrls = $r.SiteUrls
            $row.UnsupportedSoapCalls = $r.UnsupportedSoapCalls
            $row.UnsupportedSoapCallsCount = $r.UnsupportedSoapCallsCount
            $row.UnsupportedDataConnectionTypes = $r.UnsupportedDataConnectionTypes
            $row.UnsupportedDataConnectionInstances = $r.UnsupportedDataConnectionInstances
            $row.ManagedCode = $r.ManagedCode
            $row.ManagedCodeState = $r.ManagedCodeState
            $row.Mode = $r.Mode
            $row.ProductVersion = $r.ProductVersion             
                        
            
            $dirName =$URL.Substring(0, $URL.LastIndexOf("/"))
            $infoPathDataRows = $infoPathFileInfoData | ? { $_.DirName -eq $dirName }
           
            if ($infoPathDataRows -ne $null)
            {
                $logRow = $infoPathDataRows[0];
                $site = Get-SPSite -Identity $logRow.SiteId                
                                    
                #$author,$itemCount,$lastModifiedDate,$owners,$admins, $listName, $ctypeLists = Get-MetaData $URL $site
                $itemCount,$lastModifiedDate,$listName, $ctypeLists = Get-MetaData $URL $site

                $row.ListName = $listName;
                $row.ItemCount = $itemCount
                $row.LastModifiedDate = $lastModifiedDate
                <#$row.Author = $author
                $row.SiteOwners = $owners
                $row.SiteAdmins = $admins  

                if ($isXSNFileDownloaded -eq $false)
                {
                    $webAppURL = $site.WebApplication.Url
                    $infoPathURL = $webAppURL + $URL
                    
                    Download-InfoPathForm $infoPathURL $site $usageInfoPathFolder $logRow.Id

                    $isXSNFileDownloaded = $true
                }#>                 
            } #if ($logData -ne $null)
            
                   
            $table.Rows.Add($row)

            # Add content type related lists to the csv file
            foreach($ctype in $ctypeLists)
            {
                $row = $table.NewRow()
                
                $row.Urn = $r.Urn
                $row.ListName = $ctype.ListName;
                $row.ItemCount = $ctype.ItemCount
                $row.LastModifiedDate = $ctype.ModifiedDate
                <#$row.Author = $ctype.Author
                $row.SiteOwners = $owners
                $row.SiteAdmins = $admins#>
                  
                $table.Rows.Add($row) 
            } # foreach($ctype in $ctypeLists)
        } #foreach($URL in $URLs)
               
    }
}

Get-Data

$table | export-csv $outputFile


