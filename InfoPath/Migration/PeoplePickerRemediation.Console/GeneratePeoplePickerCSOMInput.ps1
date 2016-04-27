$inputFilePath = "C:\work\InfoPath\peoplepicker\InfoPathForms_PeoplePicker.csv"
$outFilePath = "C:\work\InfoPath\peoplepicker\PeoplePickerCSOMInputFile.csv"
$siteSearch = "/sites/"
$personalSearch = "/personal/"
$listsSearch = "/Lists/"
$formsSearch = "/Forms/"
$reportOutput = @()

$rows = @(Import-CSV $inputFilePath | select SiteUrls, URLs)


function GetWebApplicationUrl($csvRows, $search)
{
    $url = ""
    
    foreach($csvRow in $csvRows)
    {
        if ($csvRow.SiteUrls.Contains($search))
        {
            $url = $csvRow.SiteUrls
            $i = 1
            $findIndex = 3 

            ForEach ($match in ($url | select-String "/" -allMatches).matches)
            {
                $index = $match.Index
                if ( $i -eq $findIndex)
                {
                    $url = $url.substring(0,$index)
                }
                $i++
            }

            break
        }
    }

    return $url
}

Function GetWebUrl($siteSuffix)
{
    if ($siteSuffix.Contains("personal/"))
    {
        $webUrl = $myWebAppUrl+"/"+$siteSuffix
    }
    else
    {
        $webUrl = $webAppUrl+"/"+$siteSuffix
    }

    return $webUrl
}

$webAppUrl = GetWebApplicationUrl $rows $siteSearch
$myWebAppUrl = GetWebApplicationUrl $rows $personalSearch 

foreach($row in $rows)
{
    $formUrls = $row.URLs

    if ($formUrls.Contains(";"))
    {
        $formUrls = @($formUrls.Split(";"))
    }

    foreach($formUrl in $formUrls)
    {
        if ($formUrl.Contains("/Workflows/") -or $formUrl.Contains("/_catalogs/") -or $formUrl.Contains("/FormServerTemplates/")  -or $formUrl.Contains("/SiteAssets/"))
        {
            continue
        }
        
        if ($formUrl.Contains($listsSearch))
        {
            $lstIdx = $formUrl.IndexOf($listsSearch)
            $siteSuffix = $formUrl.SubString(0, $lstIdx)
            $listSuffix = $formUrl.SubString($lstIdx + $listsSearch.Length)
            $listName = $listSuffix.SubString(0, $listSuffix.IndexOf("/"))
            $webUrl = GetWebUrl $siteSuffix            
            
            $rowObject = New-Object -TypeName PSObject
            $rowObject | Add-Member -MemberType NoteProperty -Name WebUrl -Value $webUrl
            $rowObject | Add-Member -MemberType NoteProperty -Name ListName -Value $listName 
            $reportOutput += $rowObject     
        }
        elseif($formUrl.Contains($formsSearch))
        {
            $frmIdx = $formUrl.IndexOf($formsSearch)
            $siteSuffix = $formUrl.SubString(0, $frmIdx)
            $lastIdx = $siteSuffix.LastIndexOf("/")
            $formSuffix = $siteSuffix.SubString($lastIdx+1)
            $siteSuffix = $siteSuffix.SubString(0, $lastIdx)
            $webUrl = GetWebUrl $siteSuffix           

            $rowObject = New-Object -TypeName PSObject
            $rowObject | Add-Member -MemberType NoteProperty -Name WebUrl -Value $webUrl
            $rowObject | Add-Member -MemberType NoteProperty -Name ListName -Value $formSuffix 
            $reportOutput += $rowObject     
        }
        else
        {
            $rowObject = New-Object -TypeName PSObject
            $rowObject | Add-Member -MemberType NoteProperty -Name WebUrl -Value $row.SiteUrls
            $rowObject | Add-Member -MemberType NoteProperty -Name ListName -Value $formUrl
            $reportOutput += $rowObject    
        }
    }

}

$reportOutput | export-csv -notypeinformation $outFilePath
