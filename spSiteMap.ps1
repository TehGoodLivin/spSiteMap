#
#    DATE: 02 Nov 2021
#    UPDATED: 23 Aug 2021
#    
#    MIT License
#    Copyright (c) 2021 Austin Livengood
#    Permission is hereby granted, free of charge, to any person obtaining a copy
#    of this software and associated documentation files (the "Software"), to deal
#    in the Software without restriction, including without limitation the rights
#    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
#    copies of the Software, and to permit persons to whom the Software is
#    furnished to do so, subject to the following conditions:
#    The above copyright notice and this permission notice shall be included in all
#    copies or substantial portions of the Software.
#    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
#    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
#    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
#    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
#    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
#    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
#    SOFTWARE.
#

# CHANGABLE VARIABLES
$sitePath = "https://usaf.dps.mil/sites/52msg/CS/SCX/IAO/" # SITE PATH
$reportPath = "C:\users\$env:USERNAME\Desktop\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_SiteMapResults.csv" # REPORT PATH (DEFAULT IS TO DESK
$results = @() # RESULTS

Connect-PnPOnline -Url $sitePath -UseWebLogin # CONNECT TO SPO

$siteInfo = Get-PnPWeb -Includes Created | select Title, ServerRelativeUrl, Url, Created, Description
$siteLists = Get-PnPList | Where-Object {$_.Hidden -eq $false}
$subSites = Get-PnPSubWeb -Recurse | select Title, ServerRelativeUrl, Url, Created, Description

$siteListCount = @()
$siteItemCount = 0
foreach ($list in $subSiteLists) {
    $siteListCount += $list
    $siteItemCount = $siteItemCount + $list.ItemCount
}

# GET PARENT SITE INFO AND LIST COUNT
$results += New-Object PSObject -Property @{
    Title = $siteInfo.Title
    ItemCount = $siteItemCount
    ListCount = $siteListCount.Count
    ServerRelativeUrl = $siteInfo.ServerRelativeUrl
    Description = $siteInfo.Description
    Created = $siteInfo.Created
}

foreach ($site in $subSites) {
    Connect-PnPOnline -Url $site.Url -UseWebLogin # CONNECT TO SPO SUBSITE
    $subSiteLists = Get-PnPList | Where-Object {$_.Hidden -eq $false}

    $subSiteListCount = @()
    $subSiteItemCount = 0
    foreach ($list in $subSiteLists) {
        $subSiteListCount += $list
        $siteListCount += $list
        $subSiteItemCount = $subSiteItemCount + $list.ItemCount
        $siteItemCount = $siteItemCount + $list.ItemCount
    }

    $results += New-Object PSObject -Property @{
        Title = $site.Title
        ListCount = $subSiteListCount.Count
        ItemCount = $subSiteItemCount
        ServerRelativeUrl = $site.ServerRelativeUrl
        Description = $site.Description
        Created = $site.Created
    }
}

# GET TOTAL COUNTS
$results += New-Object PSObject -Property @{
    Title = "Total"
    ListCount = $siteListCount.Count
    ItemCount = $siteItemCount
    ServerRelativeUrl = $subSites.Count + 1
    Description = ""
    Created = ""
}
$results | Select-Object "Title", "ServerRelativeUrl", "ListCount", "ItemCount", "Description", "Created" | Export-Csv -Path $reportPath -NoTypeInformation
