$xmlapi1BaseUrl = 'https://boardgamegeek.com/xmlapi/'
$xmlapi2BaseUrl = 'https://boardgamegeek.com/xmlapi2/'
$requestThrottle = 2
$username = 'AggieEngineer2k'

enum Subtype {
    boardgame
    boardgameexpansion
    boardgameaccessory
}

<#
.SYNOPSIS
   Fetch a user's collection
#>
function Fetch-BoardGameGeekCollection {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    Param
    (
        # Username to fetch the collection for
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $username,

        # Subtype to filter on
        [Parameter(Mandatory=$false)]
        [Subtype]$subtype = [Subtype]::boardgame,

        # Subtype to exclude
        [Parameter(Mandatory=$false)]
        [Subtype]$excludesubtype,

        # Filter on owned
        [Parameter(Mandatory=$false)]
        [bool]$own = $null
    )

    $response = $null
    try {
        $url = "{0}collection?username={1}&subtype={2}" -f 
            $script:baseUrl, 
            [uri]::EscapeDataString($username),
            [uri]::EscapeDataString($subtype),
            [uri]::EscapeDataString($excludesubtype)
        if($excludesubtype -ne $null) {
            $url = "{0}&excludesubtype={1}" -f
                $url,
                [uri]::EscapeDataString($excludesubtype)
        }
        if($own -ne $null) {
            $url = "{0}&own={1}" -f
                $url,
                $(if ($own -eq $true) {'1'} else {'0'})
        }
        $request = { Invoke-WebRequest -Method Get -Uri $url -UseBasicParsing }
        $response = Invoke-Command -ScriptBlock $request
        while ($response.StatusCode -eq 202) {
            sleep -Seconds $script:requestThrottle
            $response = Invoke-Command -ScriptBlock $request
    }
    } catch [System.Net.WebException] {      
        Write-Error $_.Exception.Message
    } catch {
        Write-Warning $_.Exception.GetType().FullName
        Write-Warning $_.Exception -ForegroundColor
    }

    # Internal throttle to slow requests
    sleep -Seconds $script:requestThrottle

    return [xml]$response.Content
}

<#
.SYNOPSIS
   Fetch a user's collection
#>
function Fetch-BoardGameGeekBoardgame {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    Param
    (
        # Username to fetch the collection for
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        [int]$gameid
    )

    $response = $null
    try {
        $url = "{0}boardgame/{1}" -f 
            $script:xmlapi1baseUrl, 
            $gameid
        $request = { Invoke-WebRequest -Method Get -Uri $url -UseBasicParsing }
        $response = Invoke-Command -ScriptBlock $request
        while ($response.StatusCode -eq 202) {
            sleep -Seconds $script:requestThrottle
            $response = Invoke-Command -ScriptBlock $request
        }
    } catch [System.Net.WebException] {      
        Write-Error $_.Exception.Message
    } catch {
        Write-Warning $_.Exception.GetType().FullName
        Write-Warning $_.Exception
    }

    # Internal throttle to slow requests
    sleep -Seconds $script:requestThrottle

    return [xml]$response.Content
}

<#
.SYNOPSIS
   Fetch a thing
#>
function Fetch-BoardGameGeekThing {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    Param
    (
        # Username to fetch the collection for
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $id,

        # Type to filter on
        [Parameter(Mandatory=$false)]
        [Subtype]$type = [Subtype]::boardgame
    )

    $response = $null
    try {
        $url = "{0}thing?id={1}&type={2}" -f 
            $script:baseUrl,
            [uri]::EscapeDataString($id),
            [uri]::EscapeDataString($type)
        $request = { Invoke-WebRequest -Method Get -Uri $url -UseBasicParsing }
        $response = Invoke-Command -ScriptBlock $request
        while ($response.StatusCode -eq 202 -or $response.StatusCode -eq 429) {
            sleep -Seconds $script:requestThrottle
            $response = Invoke-Command -ScriptBlock $request
    }
    } catch [System.Net.WebException] {      
        Write-Error $_.Exception.Message
    } catch {
        Write-Warning $_.Exception.GetType().FullName
        Write-Warning $_.Exception
    }

    # Internal throttle to slow requests
    sleep -Seconds $script:requestThrottle

    return [xml]$response.Content
}

Write-Progress -Activity 'Fetching owned boardgames...'
$collection = Fetch-BoardGameGeekCollection -username $username -own 1 -subtype boardgame -excludesubtype boardgameexpansion
$collection.InnerXml | Out-File -FilePath "C:\Users\Justin L. Brown\Desktop\boardgame.xml"
#[xml]$collection = Get-Content -Path 'C:\Users\Justin L. Brown\Desktop\boardgame.xml'
$ownedBoardgameObjectIds = $collection.items.item.objectid

Write-Progress -Activity 'Fetching owned expansions...'
$collection = Fetch-BoardGameGeekCollection -username $username -own 1 -subtype boardgameexpansion
$collection.InnerXml | Out-File -FilePath "C:\Users\Justin L. Brown\Desktop\boardgameexpansion.xml"
#[xml]$collection = Get-Content -Path 'C:\Users\Justin L. Brown\Desktop\boardgameexpansion.xml'
$ownedBoardgameExpansionObjectIds = $collection.items.item.objectid

Write-Progress -Activity 'Fetching owned accessories...'
$collection = Fetch-BoardGameGeekCollection -username $username -own 1 -subtype boardgameaccessory
$collection.InnerXml | Out-File -FilePath "C:\Users\Justin L. Brown\Desktop\boardgameaccessory.xml"
#[xml]$collection = Get-Content -Path 'C:\Users\Justin L. Brown\Desktop\boardgameaccessory.xml'
$ownedBoardgameAccessoryObjectIds = $collection.items.item.objectid

$knownExpansionIds = [System.Collections.ArrayList]@()
$knownAccessoryIds = [System.Collections.ArrayList]@()

$count = $ownedBoardgameObjectIds.Count
for ($i = 0; $i -lt $count; $i++) {
    Write-Progress -Activity 'Iterating Boardgames' -Status "$($i + 1) of $count" -PercentComplete (($i/$count) * 100)

    $id = $ownedBoardgameObjectIds[$i]
    $boardgame = Fetch-BoardGameGeekBoardgame -gameid $id

    Select-Xml -Xml $boardgame -XPath '/boardgames/boardgame/boardgameexpansion' | Select-Object -ExpandProperty Node | ForEach-Object { if ($ownedBoardgameExpansionObjectIds -notcontains $_.objectid) { $knownExpansionIds.Add($_.objectid) | Out-Null } }
    Select-Xml -Xml $boardgame -XPath '/boardgames/boardgame/boardgameaccessory' | Select-Object -ExpandProperty Node | ForEach-Object { if ($ownedBoardgameAccessoryObjectIds -notcontains $_.objectid) { $knownAccessoryIds.Add($_.objectid) | Out-Null } }
}

$html = [System.Text.StringBuilder]::new()
$html.Append('<html><body>') | Out-Null

$html.Append('<h1>Expansions</h1>')
$html.Append('<table>') | Out-Null
$count = $knownExpansionIds.Count
for ($i = 0; $i -lt $count; $i++) {
    $id = $knownExpansionIds[$i]
    $thing = Fetch-BoardGameGeekThing -id $id -type boardgameexpansion

    $thumbnail = (Select-Xml -Xml $thing -XPath '/items/item/thumbnail').Node.InnerText
    $name = (Select-Xml -Xml $thing -XPath '/items/item/name[@type="primary"]').Node.Value
    $url = "https://boardgamegeek.com/boardgame/$id"

    Write-Progress -Activity 'Iterating Expansions' -Status "$($i + 1) of $count" -PercentComplete (($i/$count) * 100)
    
    $html.Append('<tr>') | Out-Null
    $html.Append("<td><img src='$thumbnail'/></td>") | Out-Null
    $html.Append("<td><a href='$url'>$name</a></td>") | Out-Null
    $html.Append('</tr>') | Out-Null
}
$html.Append('</table>') | Out-Null

$html.Append('<h1>Accessories</h1>')
$html.Append('<table>') | Out-Null
$count = $knownAccessoryIds.Count
for ($i = 0; $i -lt $count; $i++) {
    $id = $knownAccessoryIds[$i]
    $thing = Fetch-BoardGameGeekThing -id $id -type boardgameaccessory

    $thumbnail = (Select-Xml -Xml $thing -XPath '/items/item/thumbnail').Node.InnerText
    $name = (Select-Xml -Xml $thing -XPath '/items/item/name[@type="primary"]').Node.Value
    $url = "https://boardgamegeek.com/boardgame/$id"

    Write-Progress -Activity 'Iterating Accessories' -Status "$($i + 1) of $count" -PercentComplete (($i/$count) * 100)
    
    $html.Append('<tr>') | Out-Null
    $html.Append("<td><img src='$thumbnail'/></td>") | Out-Null
    $html.Append("<td><a href='$url'>$name</a></td>") | Out-Null
    $html.Append('</tr>') | Out-Null
}
$html.Append('</table>') | Out-Null

$html.Append('</body></html>') | Out-Null
$html.ToString() | Out-File -FilePath 'C:\Users\Justin L. Brown\Desktop\bgfomo.html'