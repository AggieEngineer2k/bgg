$baseUrl = 'https://boardgamegeek.com/xmlapi2/'
$requestThrottle = 3

enum Subtype {
    boardgame
    boardgameexpansion
    boardgameaccessory
}

<#
.SYNOPSIS
   Fetch a user's collection that they own
#>
function Fetch-BoardGameGeekCollectionOwned {
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
        [Subtype]$excludesubtype
    )

    $response = $null
    try {
        $url = "{0}collection?username={1}&subtype={2}&own=1" -f 
            $script:baseUrl, 
            [uri]::EscapeDataString($username),
            [uri]::EscapeDataString($subtype),
            [uri]::EscapeDataString($excludesubtype)
        if($excludesubtype -ne $null) {
            $url = "{0}&excludesubtype={1}" -f
                $url,
                [uri]::EscapeDataString($excludesubtype)
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
        Write-Warning $_.Exception.GetType().FullName -ForegroundColor Yellow
        Write-Warning $_.Exception -ForegroundColor Red
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
        while ($response.StatusCode -eq 202) {
            sleep -Seconds $script:requestThrottle
            $response = Invoke-Command -ScriptBlock $request
    }
    } catch [System.Net.WebException] {      
        Write-Error $_.Exception.Message
    } catch {
        Write-Warning $_.Exception.GetType().FullName -ForegroundColor Yellow
        Write-Warning $_.Exception -ForegroundColor Red
    }

    # Internal throttle to slow requests
    sleep -Seconds $script:requestThrottle

    return [xml]$response.Content
}

<#
.SYNOPSIS
   Fetch a thing
#>
function Fetch-BoardGameGeekFamily {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    Param
    (
        # Username to fetch the collection for
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $id
    )

    $response = $null
    try {
        $url = "{0}family?id={1}" -f 
            $script:baseUrl, 
            [uri]::EscapeDataString($id)
        $request = { Invoke-WebRequest -Method Get -Uri $url -UseBasicParsing }
        $response = Invoke-Command -ScriptBlock $request
        while ($response.StatusCode -eq 202) {
            sleep -Seconds $script:requestThrottle
            $response = Invoke-Command -ScriptBlock $request
    }
    } catch [System.Net.WebException] {      
        Write-Error $_.Exception.Message
    } catch {
        Write-Warning $_.Exception.GetType().FullName -ForegroundColor Yellow
        Write-Warning $_.Exception -ForegroundColor Red
    }

    # Internal throttle to slow requests
    sleep -Seconds $script:requestThrottle

    return [xml]$response.Content
}

#$collection = Fetch-BoardGameGeekCollectionOwned -username 'AggieEngineer2k' -subtype boardgame -excludesubtype boardgameexpansion
#$collection.InnerXml | Out-File -FilePath "C:\Users\Justin L. Brown\Desktop\boardgame.xml"
[xml]$collection = Get-Content -Path 'C:\Users\Justin L. Brown\Desktop\boardgame.xml'
$ownedBoardgameObjectIds = $collection.items.item.objectid

#$collection = Fetch-BoardGameGeekCollectionOwned -username 'AggieEngineer2k' -subtype boardgameexpansion
#$collection.InnerXml | Out-File -FilePath "C:\Users\Justin L. Brown\Desktop\boardgameexpansion.xml"
[xml]$collection = Get-Content -Path 'C:\Users\Justin L. Brown\Desktop\boardgameexpansion.xml'
$ownedBoardgameExpansionObjectIds = $collection.items.item.objectid

#$collection = Fetch-BoardGameGeekCollectionOwned -username 'AggieEngineer2k' -subtype boardgameaccessory
#$collection.InnerXml | Out-File -FilePath "C:\Users\Justin L. Brown\Desktop\boardgameaccessory.xml"
#[xml]$collection = Get-Content -Path 'C:\Users\Justin L. Brown\Desktop\boardgameaccessory.xml'
#$ownedBoardgameAccessoryObjectIds = $collection.items.item.objectid

$expansionIds = [System.Collections.ArrayList]@()

$count = $ownedBoardgameObjectIds.Count
for ($i = 0; $i -lt $count; $i++) {
    $id = $ownedBoardgameObjectIds[$i]
    $thing = Fetch-BoardGameGeekThing -id $id

    Write-Progress -Activity 'Iterating Boardgames' -Status "$($i + 1) of $count" -PercentComplete (($i/$count) * 100)

    Select-Xml -Xml $thing -XPath '/items/item/link[@type="boardgameexpansion"]' | Select-Object -ExpandProperty Node | ForEach-Object { if ($ownedBoardgameExpansionObjectIds -notcontains $_.Id) { $expansionIds.Add($_.Id) | Out-Null } }
}

$html = [System.Text.StringBuilder]::new()
$html.Append('<html><body>') | Out-Null

$html.Append('<table>') | Out-Null
$count = $expansionIds.Count
for ($i = 0; $i -lt $count; $i++) {
    $id = $expansionIds[$i]
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

$html.Append('</body></html>') | Out-Null
$html.ToString() | Out-File -FilePath 'C:\Users\Justin L. Brown\Desktop\expansions.html'