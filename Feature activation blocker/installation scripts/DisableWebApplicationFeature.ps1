
$urls = @("<web application>","<web application>")

function DisableFeature($url)
{
    Disable-SPFeature –Identity d073413c-3bd4-4fcc-a0ff-61e0e8043843 -URL $url -Confirm:$false
    Write-Host "Deactivated feature d073413c-3bd4-4fcc-a0ff-61e0e8043843 on web application $url"
}

foreach($url in $urls)
{
    DisableFeature $url
}
