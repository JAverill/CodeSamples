## Google Custom API Search Script
## Version 1.0 - Author: John Averill (Written for private client.)

## Overview: A script that, when used to call a complex Custom Search API from Google, will attempt to capture all textual data from each of the top 50 resulting webpages to any user query, parse said HTML data into legible format and output it to single document. The script is to be used for opposition research for a political campaign, to seek websites with specific text phrases beyond the capacity of a standard search engine search.

## Usage: Script is merely executed from a PowerShell Terminal, prompts for query and executes. Logs output to C:\TESTFOLDER\lastsearchresults.csv.

## Add Assemblies
Add-Type -AssemblyName System.Web

## Set Error Action Preference, as errors are compensated for in script.
$erroractionpreference = 'SilentlyContinue'

## Set Global Variables and Creds
$global:GoogleCSEAPIKey = "CSE API Key Here"
$global:GoogleCSEIdentifier = "CSE Identifier"

## Get Query String Function, converts user input to HTTP URL-Encoded query.
Function Get-QueryString {
    param([string[]] $Query)
    $Global:QueryString = ($Query | %{
        [Web.HTTPUtility]::UrlEncode($_)}) -join '+'
}

## Scrape Now Function, passes query to Google REST API to collect top 50 results within custom search parameters. Note, Google advises it is most efficient to limit output of queries to 10 result batches using this method with a free account and iterate through as needed.
Function Scrape-Now {
    $global:searchstring = Read-Host -Prompt 'Enter Search String'
    Get-QueryString -Query "$global:searchstring"
    $global:results = [system.collections.arraylist]@{}
    $global:uri = "https://customsearch.googleapis.com/customsearch/v1?key=$global:GoogleCSEAPIKey&cx=$global:GoogleCSEIdentifier&q=$global:QueryString&start=1&num=10"
    $global:results += Invoke-RestMethod -URI $global:uri
    $global:uri = "https://customsearch.googleapis.com/customsearch/v1?key=$global:GoogleCSEAPIKey&cx=$global:GoogleCSEIdentifier&q=$global:QueryString&start=11&num=10"
    $global:results += Invoke-RestMethod -URI $global:uri
    $global:uri = "https://customsearch.googleapis.com/customsearch/v1?key=$global:GoogleCSEAPIKey&cx=$global:GoogleCSEIdentifier&q=$global:QueryString&start=21&num=10"
    $global:results += Invoke-RestMethod -URI $global:uri
    $global:uri = "https://customsearch.googleapis.com/customsearch/v1?key=$global:GoogleCSEAPIKey&cx=$global:GoogleCSEIdentifier&q=$global:QueryString&start=31&num=10"
    $global:results += Invoke-RestMethod -URI $global:uri
    $global:uri = "https://customsearch.googleapis.com/customsearch/v1?key=$global:GoogleCSEAPIKey&cx=$global:GoogleCSEIdentifier&q=$global:QueryString&start=41&num=10"
    $global:results += Invoke-RestMethod -URI $global:uri
    $global:totalresults = $global:results | Select-Object -ExpandProperty Items
}

## Parse Page Function, which will be executed on each of the 50 Custom Search Results to Parse the raw HTML of that page - only for text/paragraph blocks - into a legible format.
Function Parse-Page{
    $global:uri = $global:result.Link
    $global:sitedata = Invoke-WebRequest -URI $global:uri
    IF(!$global:sitedata.RawContent){$global:output = 'No Site Data Pulled'}
    ELSE{
        $global:sitedata.RawContent | Out-File C:\TESTFOLDER\temp.txt -Force
        $global:import = Get-Content C:\TESTFOLDER\temp.txt
        Remove-Item -Path C:\TESTFOLDER\temp.txt -Force
        IF(!$global:import){$global:output = 'No Site Data Pulled'}
        ELSE{
            $global:y = $global:import | Select-String -Pattern '<p>'
            $global:b = $global:y -Replace '<[^>]+>',''
            $global:siteresults = $global:b -Replace '&nbsp'," "
            $global:trim = $global:siteresults.trim()
            IF(!$global:trim){[string]$global:output = 'Site Data Not Parsed'}
            ELSEIF($global:trim -like "window._*"){$global:output = 'Site Data Not Parsed'}
            ELSE{
                [string]$global:output = $global:trim -replace '\s+',' '
            }
        }
    }
}

## Log Results Function, which will be used to log the scrape of the top 50 results to a CSV for easy skimming.
Function Log-Results{
ForEach($global:result in $global:totalresults){
    Parse-Page
    $global:logging = New-Object PSObject
    $global:logging | Add-Member NoteProperty Title $global:result.Title
    $global:logging | Add-Member NoteProperty Snippet $global:result.Snippet
    $global:logging | Add-Member NoteProperty Link $global:result.Link
    $global:logging | Add-Member NoteProperty DetailedData $global:output
    $global:logging | Export-CSV "C:\TESTFOLDER\lastsearchresults.csv" -NoTypeInformation -Append
    $global:output = 'Site Data Not Parsed'}
}

## Execute Functions.
Scrape-Now
Log-Results