<#-----------------------------------------------------------------------------------------#
Author: Uzair Ansari

Function: This script will download the KB articles from Microsoft update catalog
          Input list of KB articles need to be provided

Date: 25th June 2018

Version 1.1 : Added OS filter

Version: 1.0

#-----------------------------------------------------------------------------------------#>


<#-----------------------------------------------------------------------------------------#
NOTES:
1) By default this script will download all the KB articles present on the update
   catalog website. If you want to download only specific updates like 'x64' or
   only 'Cumulative' updates then you need to remove the # preceding the type you want
   to download. For example if you want to download only Cumulative updates for x64 bit
   machines then remove the # preceding "#$Platform = 'x64'" and #$Type = 'Cumulative'"


2) Kindly provide the path where you wish to download the patches in $parentpath variable.
   Kindly provide the text file path where KB article details are saved in $kblist variable.
   KB details should be saved one below the other in text file as:
   
   KB123456
   KB567890
   
#-----------------------------------------------------------------------------------------#>
Clear-Variable -Name OS | Out-Null
Clear-Variable -Name Type | Out-Null
Clear-Variable -Name Platform | Out-Null


<#-----------------------------------------------------------------------------------------#
Remove the # before the filter you want to apply.
#-----------------------------------------------------------------------------------------#>

$Platform = 'x64'
#$Platform = 'x86'
#$Platform = 'arm64'


#$Type = 'Cumulative'
$Type = 'Delta'


#$OS = "Windows Server 2016"
$OS = 'Windows 10'




$ParentPath="$env:USERPROFILE"

$Kblist=Get-Content 'D:\KBlist.txt'



function Find-AssetUri {
<#
    Get Uri of Microsoft download asset specified by its GUID
#>
    [CmdletBinding()]
    Param(
        [String[]]
            [Parameter(
                Mandatory,
                Position = 0,
                ValueFromPipeline,
                ValueFromPipelineByPropertyName,
                HelpMessage = "GUID of Microsoft download asset."
            )]
            [Alias( 'Id' )]
        $GUID
    )

    Begin {
        $updateCatalogDownloadLink = 'http://www.catalog.update.microsoft.com/DownloadDialog.aspx'

        $assetUriPattern =  "https?://download\.windowsupdate\.com\/[^ \'\""]+"
        $postBodyTemplate = '"size": 0,  "uidInfo": "{0}",  "updateID": "{0}"' -replace ' ', ''
    }

    Process {
        foreach ($oneGUID in $GUID) {
            Write-Verbose "Download description of asset $oneGUID"
            $postBody = @{ updateIDs = "[{$( $postBodyTemplate -f $oneGUID )}]" }

            if (    ( Invoke-WebRequest -Uri $updateCatalogDownloadLink -Method Post -Body $postBody
                    ).Content -match $assetUriPattern
            ) {
                $Matches[0]
            }
        }
    }

    End {}
}



foreach ($KB in $Kblist)
{

    $oneUri="https://www.catalog.update.microsoft.com/Search.aspx?q=$KB"


    #Type filter
    If ($Type -ne $null -and $Platform -eq $null -and $OS -eq $null)
        {
         $Links=(Invoke-WebRequest -Uri $oneUri).Links | Where-Object id -like '*_link' | Select-Object @{Name = 'GUID';Expression = { $_.Id -replace '_link', '' }}, @{Name = 'platform';Expression = {if( $_.innerText -match '\b(x86|x64|arm64)\b') {$Matches[1].ToLower()}}}, @{Name = 'type';Expression = {if( $_.innerText -match '\b(Cumulative|Delta)\b') {$Matches[1].ToLower()}}}, class, innertext, href | Where-Object {($_.Type -eq $Type)}
        }


    #Platform filter
    elseIf ($Platform -ne $null -and $Type -eq $null -and $OS -eq $null)
        {
         $Links=(Invoke-WebRequest -Uri $oneUri).Links | Where-Object id -like '*_link' | Select-Object @{Name = 'GUID';Expression = { $_.Id -replace '_link', '' }}, @{Name = 'platform';Expression = {if( $_.innerText -match '\b(x86|x64|arm64)\b') {$Matches[1].ToLower()}}}, @{Name = 'type';Expression = {if( $_.innerText -match '\b(Cumulative|Delta)\b') {$Matches[1].ToLower()}}}, class, innertext, href | Where-Object {($_.Platform -eq $Platform)}
        }


    #OS filter
    elseIf ($OS -ne $null -and $Type -eq $null -and $Platform -eq $null)
        {
         $Links=(Invoke-WebRequest -Uri $oneUri).Links | Where-Object id -like '*_link' | Select-Object @{Name = 'GUID';Expression = { $_.Id -replace '_link', '' }}, @{Name = 'platform';Expression = {if( $_.innerText -match '\b(x86|x64|arm64)\b') {$Matches[1].ToLower()}}}, @{Name = 'type';Expression = {if( $_.innerText -match '\b(Cumulative|Delta)\b') {$Matches[1].ToLower()}}}, class, innertext, href | Where-Object {($_.innertext -like "*$OS*")}
        }


    #OS and Platform filter
    elseIf ($OS -ne $null -and $Platform -ne $null -and $Type -eq $null)
        {
         $Links=(Invoke-WebRequest -Uri $oneUri).Links | Where-Object id -like '*_link' | Select-Object @{Name = 'GUID';Expression = { $_.Id -replace '_link', '' }}, @{Name = 'platform';Expression = {if( $_.innerText -match '\b(x86|x64|arm64)\b') {$Matches[1].ToLower()}}}, @{Name = 'type';Expression = {if( $_.innerText -match '\b(Cumulative|Delta)\b') {$Matches[1].ToLower()}}}, class, innertext, href | Where-Object {($_.innertext -like "*$OS*") -and ($_.Platform -eq $Platform)}
        }


    #OS and Type filter
    elseIf ($OS -ne $null -and $Type -ne $null -and $Platform -eq $null)
        {
         $Links=(Invoke-WebRequest -Uri $oneUri).Links | Where-Object id -like '*_link' | Select-Object @{Name = 'GUID';Expression = { $_.Id -replace '_link', '' }}, @{Name = 'platform';Expression = {if( $_.innerText -match '\b(x86|x64|arm64)\b') {$Matches[1].ToLower()}}}, @{Name = 'type';Expression = {if( $_.innerText -match '\b(Cumulative|Delta)\b') {$Matches[1].ToLower()}}}, class, innertext, href | Where-Object {($_.innertext -like "*$OS*") -and ($_.Type -eq $Type)}
        }


    #Type and Platform filter
    elseIF ($Platform -ne $null -and $Type -ne $null -and $OS -eq $null)
        {
         $Links=(Invoke-WebRequest -Uri $oneUri).Links | Where-Object id -like '*_link' | Select-Object @{Name = 'GUID';Expression = { $_.Id -replace '_link', '' }}, @{Name = 'platform';Expression = {if( $_.innerText -match '\b(x86|x64|arm64)\b') {$Matches[1].ToLower()}}}, @{Name = 'type';Expression = {if( $_.innerText -match '\b(Cumulative|Delta)\b') {$Matches[1].ToLower()}}}, class, innertext, href | Where-Object {($_.Platform -eq $Platform) -and ($_.Type -eq $Type)}
        }


    #No Filter
    elseIf ($Platform -eq $null -and $Type -eq $null -and $OS -eq $null)
        {
         $Links=(Invoke-WebRequest -Uri $oneUri).Links | Where-Object id -like '*_link' | Select-Object @{Name = 'GUID';Expression = { $_.Id -replace '_link', '' }}, @{Name = 'platform';Expression = {if( $_.innerText -match '\b(x86|x64|arm64)\b') {$Matches[1].ToLower()}}}, @{Name = 'type';Expression = {if( $_.innerText -match '\b(Cumulative|Delta)\b') {$Matches[1].ToLower()}}}, class, innertext, href
        }


    #All Filter
    elseIf ($Platform -ne $null -and $Type -ne $null -and $OS -ne $null)
        {
         $Links=(Invoke-WebRequest -Uri $oneUri).Links | Where-Object id -like '*_link' | Select-Object @{Name = 'GUID';Expression = { $_.Id -replace '_link', '' }}, @{Name = 'platform';Expression = {if( $_.innerText -match '\b(x86|x64|arm64)\b') {$Matches[1].ToLower()}}}, @{Name = 'type';Expression = {if( $_.innerText -match '\b(Cumulative|Delta)\b') {$Matches[1].ToLower()}}}, class, innertext, href | Where-Object {($_.Platform -eq $Platform) -and ($_.Type -eq $Type) -and ($_.innertext -like "*$OS*")}
        }



    foreach ($i in $Links)
        {
         $FolderNameTemp=$i.innerText
         $foldername=$FolderNameTemp.TrimEnd(" ")
         $DownloadLink=$i.GUID | Find-AssetUri
         $FileNameTemp=($DownloadLink -split "/")
         $FileName=$FileNameTemp[-1]

         New-Item -ItemType Directory -Path "$ParentPath\$KB" -ErrorAction SilentlyContinue
         New-Item -ItemType Directory -Path "$ParentPath\$KB\$foldername"
         $DownloadPath="$ParentPath\$KB\$foldername\$FileName"


         Invoke-WebRequest $DownloadLink -OutFile $DownloadPath
        }

}

