
$urls = @("<web application>","<web application>")

function EnableFeature($url)
{
    Enable-SPFeature –Identity d073413c-3bd4-4fcc-a0ff-61e0e8043843 -URL $url
    Write-Host "Activated feature d073413c-3bd4-4fcc-a0ff-61e0e8043843 on web application $url"
}

foreach($url in $urls)
{
    EnableFeature $url
}
