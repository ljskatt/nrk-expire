param (
    [Parameter()]
    [switch]
    $CalendarReport,

    [Parameter()]
    [switch]
    $ExcelReport
)

$time_now = Get-Date
$warn_expire = (Get-Date).AddMonths(12)
$root_location = Get-Location
$cache_dir = "$root_location\cache-nrk"
$excel_file = "$root_location\near-expire-nrk.xlsx"
$calendar_file = "$root_location\near-expire-nrk.ics"
$excel_values = @()
$ProgressPreference = 'SilentlyContinue'
$requests_cached = 0
$requests_uncached = 0
$processed_items = 0

if ((!($CalendarReport)) -and (!($ExcelReport))) {
    Write-Host -BackgroundColor "Red" -ForegroundColor "Black" -Object " -CalendarReport and/or -ExcelReport is not specified " -NoNewline; Write-Host -ForegroundColor "DarkGray" -Object "|"
    exit
}

if ($ExcelReport) {
    if (!(Get-Module -Name ImportExcel)) {
        Write-Host -BackgroundColor "Red" -ForegroundColor "Black" -Object " ImportExcel is not installed, please install before running this script " -NoNewline; Write-Host -ForegroundColor "DarkGray" -Object "|"
        exit
    }
}

if (!(Test-Path -Type "Container" -Path "$cache_dir")){
    New-Item -ItemType "Directory" -Path "$cache_dir" | Out-Null
}
if (!(Test-Path -Type "Container" -Path "$cache_dir")){
    Write-Host -BackgroundColor "Red" -ForegroundColor "Black" -Object " Could not create cache directory, exiting " -NoNewline; Write-Host -ForegroundColor "DarkGray" -Object "|"
    exit
}

if (!(Test-Path -Type "Container" -Path "$cache_dir/series")){
    New-Item -ItemType "Directory" -Path "$cache_dir/series" | Out-Null
}
if (!(Test-Path -Type "Container" -Path "$cache_dir/series")){
    Write-Host -BackgroundColor "Red" -ForegroundColor "Black" -Object " Could not create series cache directory, exiting " -NoNewline; Write-Host -ForegroundColor "DarkGray" -Object "|"
    exit
}

if (!(Test-Path -Type "Container" -Path "$cache_dir/episode")){
    New-Item -ItemType "Directory" -Path "$cache_dir/episode" | Out-Null
}
if (!(Test-Path -Type "Container" -Path "$cache_dir/episode")){
    Write-Host -BackgroundColor "Red" -ForegroundColor "Black" -Object " Could not create episode cache directory, exiting " -NoNewline; Write-Host -ForegroundColor "DarkGray" -Object "|"
    exit
}

if (!(Test-Path -Type "Container" -Path "$cache_dir/standaloneprogram")){
    New-Item -ItemType "Directory" -Path "$cache_dir/standaloneprogram" | Out-Null
}
if (!(Test-Path -Type "Container" -Path "$cache_dir/standaloneprogram")){
    Write-Host -BackgroundColor "Red" -ForegroundColor "Black" -Object " Could not create standaloneprogram cache directory, exiting " -NoNewline; Write-Host -ForegroundColor "DarkGray" -Object "|"
    exit
}

function Format-Name {
    param (
        [Parameter(Mandatory)]
        [string]
        $Name
    )
    $output = $Name
    $output = $output -replace "\?"
    $output = $output -replace ":"
    $output = $output -replace [char]0x0021 # !
    $output = $output -replace [char]0x0022 # "
    $output = $output -replace "\*"
    $output = $output -replace "/"
    $output = $output -replace '\\'
    return $output
}

$categories = ((invoke-webrequest "https://psapi.nrk.no/tv/pages").Content | ConvertFrom-Json).pageListItems._links.self.href
$requests_uncached += 1
$ProgressPreference = 'Continue'
Write-Progress -Id 0 -Activity "Requests" -Status "Uncached: $requests_uncached, Cached: $requests_cached, Processed items: $processed_items"
$ProgressPreference = 'SilentlyContinue'

$category_total = $categories.Count
$category_current = 0

foreach ($category in $categories){
    $category_current += 1
    $ProgressPreference = 'Continue'
    Write-Progress -Id 1 -Activity "Category" -Status "$category_current/$category_total" -PercentComplete (100 / $category_total * $category_current)
    $ProgressPreference = 'SilentlyContinue'
    $req = (Invoke-WebRequest -Uri "https://psapi.nrk.no$category").Content | ConvertFrom-Json
    $requests_uncached += 1

    $subcat_total = $req.sections.included.Count
    $subcat_current = 0
    foreach ($subcategory in $req.sections.included) {
        $subcat_current += 1
        $ProgressPreference = 'Continue'
        Write-Progress -Id 2 -Activity "Subcategory" -Status "$subcat_current/$subcat_total" -PercentComplete (100 / $subcat_total * $subcat_current)
        $ProgressPreference = 'SilentlyContinue'

        $progseries_total = $subcategory.plugs.Count
        $progseries_current = 0
        foreach ($series in $subcategory.plugs) {
            $ProgressPreference = 'Continue'
            $progseries_current += 1
            Write-Progress -Id 0 -Activity "Requests" -Status "Uncached: $requests_uncached, Cached: $requests_cached, Processed items: $processed_items"
            Write-Progress -Id 3 -Activity "Program / Series" -Status "$progseries_current/$progseries_total" -PercentComplete (100 / $progseries_total * $progseries_current)
            $ProgressPreference = 'SilentlyContinue'
            if ($series.targetType -eq "series") {
                $series_name = $series.displayContractContent.contentTitle
                $series_name_filtered = Format-Name -Name "$series_name"
                $series_url = $series.series._links.self.href

                if (Test-Path -PathType "Leaf" -Path "$cache_dir/series/$series_name_filtered.json"){
                    $episodes_raw = Get-Content -Path "$cache_dir/series/$series_name_filtered.json" | ConvertFrom-Json
                    $requests_cached += 1
                }
                else {
                    Invoke-WebRequest -Uri "https://psapi.nrk.no/tv/catalog$series_url" -OutFile "$cache_dir/series/$series_name_filtered.json"
                    $requests_uncached += 1
                    if (Test-Path -Type "Leaf" -Path "$cache_dir/series/$series_name_filtered.json") {
                        $episodes_raw = Get-Content -Path "$cache_dir/series/$series_name_filtered.json" | ConvertFrom-Json
                    }
                    else {
                        Write-Host -BackgroundColor "Red" -ForegroundColor "Black" -Object " Error downloading $series_name_filtered " -NoNewline; Write-Host -ForegroundColor "DarkGray" -Object "|"
                    }
                }

                $processed_items += $episodes_raw._embedded.seasons._embedded.episodes.Count
                foreach ($episode in $episodes_raw._embedded.seasons._embedded.episodes) {
                    if ($episode.usageRights.to.date){
                        $expire_date_value = Get-Date -Date $episode.usageRights.to.date
                        $episode_id = $episode.prfId
                        if ($time_now -gt $expire_date_value) {
                            # Expired
                        }
                        elseif ($warn_expire -gt $expire_date_value) {
                            # Available less than a year
                            $check_excel = $null
                            $check_excel = $excel_values | Where-Object {($_.Name -eq $series_name) -and ($_.Episode -eq $episode_id)}
                            if (!($check_excel.Name)) {
                                $hash = [ordered]@{
                                    'Name' = $series_name
                                    'URL' = "https://tv.nrk.no/episode$episode_id"
                                    'Type' = $series.targetType
                                    'Date' = (Get-Date -Format "yyyy-MM-dd" -Date $episode.usageRights.to.date)
                                    'Episode' = $episode_id
                                }
                                $excel_values += New-Object -TypeName "PSObject" -Property $hash
                            }
                        }
                    }
                }
            }
            elseif ($series.targetType -eq "episode") {
                $processed_items += 1
                $series_name = $series.displayContractContent.contentTitle
                $series_name_filtered = Format-Name -Name "$series_name"
                $series_url = $series.episode._links.self.href

                if (Test-Path -PathType "Leaf" -Path "$cache_dir/episode/$series_name_filtered.json"){
                    $episodes_raw = Get-Content -Path "$cache_dir/episode/$series_name_filtered.json" | ConvertFrom-Json
                    $requests_cached += 1
                }
                else {
                    Invoke-WebRequest -Uri "https://psapi.nrk.no/tv/catalog$series_url" -OutFile "$cache_dir/episode/$series_name_filtered.json"
                    $requests_uncached += 1
                    if (Test-Path -Type "Leaf" -Path "$cache_dir/episode/$series_name_filtered.json") {
                        $episodes_raw = Get-Content -Path "$cache_dir/episode/$series_name_filtered.json" | ConvertFrom-Json
                    }
                    else {
                        Write-Host -BackgroundColor "Red" -ForegroundColor "Black" -Object " Error downloading $series_name_filtered " -NoNewline; Write-Host -ForegroundColor "DarkGray" -Object "|"
                    }
                }

                $date = $episodes_raw.moreInformation.usageRights.to
                $expire_date_value = Get-Date -Date $date.date
                $expire_date_display = Get-Date -Format "yyyy-MM-dd" -Date $date.date

                if ($time_now -gt $expire_date_value) {
                    # Expired
                }
                elseif ($warn_expire -gt $expire_date_value) {
                    # Available less than a year
                    $check_excel = $null
                    $check_excel = $excel_values | Where-Object {($_.Name -eq "$series_name") -and ($_.Date -eq "$expire_date_display")}
                    if (!($check_excel.Name)) {
                        $hash = [ordered]@{
                            'Name' = $series_name
                            'URL' = "https://tv.nrk.no$series_url"
                            'Type' = $series.targetType
                            'Date' = $expire_date_display
                            'Episode' = ''
                        }
                        $excel_values += New-Object -TypeName "PSObject" -Property $hash
                    }
                }
            }
            elseif ($series.targetType -eq "standaloneProgram") {
                $processed_items += 1
                $series_name = $series.displayContractContent.contentTitle
                $series_name_filtered = Format-Name -Name "$series_name"
                $series_url = $series.standaloneProgram._links.self.href

                if (Test-Path -Type "Leaf" -Path "$cache_dir/standaloneprogram/$series_name_filtered.json"){
                    $episodes_raw = Get-Content -Path "$cache_dir/standaloneprogram/$series_name_filtered.json" | ConvertFrom-Json
                    $requests_cached += 1
                }
                else {
                    Invoke-WebRequest -Uri "https://psapi.nrk.no/tv/catalog$series_url" -OutFile "$cache_dir/standaloneprogram/$series_name_filtered.json"
                    $requests_uncached += 1
                    if (Test-Path -Type "Leaf" -Path "$cache_dir/standaloneprogram/$series_name_filtered.json") {
                        $episodes_raw = Get-Content -Path "$cache_dir/standaloneprogram/$series_name_filtered.json" | ConvertFrom-Json
                    }
                    else {
                        Write-Host -BackgroundColor "Red" -ForegroundColor "Black" -Object " Error downloading $series_name_filtered " -NoNewline; Write-Host -ForegroundColor "DarkGray" -Object "|"
                    }
                }

                $date = $episodes_raw.moreInformation.usageRights.to
                $expire_date_value = Get-Date -Date $date.date
                $expire_date_display = Get-Date -Format "yyyy-MM-dd" -Date $date.date

                if ($time_now -gt $expire_date_value) {
                    # Expired
                }
                elseif ($warn_expire -gt $expire_date_value) {
                    # Available less than a year
                    $check_excel = $null
                    $check_excel = $excel_values | Where-Object {($_.Name -eq "$series_name") -and ($_.Date -eq "$expire_date_display")}
                    if (!($check_excel.Name)) {
                        $hash = [ordered]@{
                            'Name' = $series_name
                            'URL' = "https://tv.nrk.no$series_url"
                            'Type' = $series.targetType
                            'Date' = $expire_date_display
                            'Episode' = ''
                        }
                        $excel_values += New-Object -TypeName "PSObject" -Property $hash
                    }
                }
            }
        }
    }
}

if ($CalendarReport) {
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("BEGIN:VCALENDAR")
    [void]$sb.AppendLine("VERSION:2.0")
    [void]$sb.AppendLine("METHOD:PUBLISH")

    $longDateFormat = "yyyyMMddTHHmmssZ"
    $calendar_values = @()

    $ProgressPreference = 'Continue'
    $excel_values_total = $excel_values.Count
    $excel_values_current = 0
    $excel_values_dropped = 0
    foreach ($row in $excel_values.GetEnumerator()) {
        $excel_values_current += 1
        $expire_date = $row.Date -replace "-" 
        $check_calendar = $null
        $check_calendar = $calendar_values | Where-Object {($_.Name -eq $row.Name) -and ($_.Date -eq "$expire_date")}
        if (!($check_calendar.Name)){
            $hash = [ordered]@{
                'Name' = $row.Name
                'Date' = $expire_date
            }
            $calendar_values += New-Object -TypeName "PSObject" -Property $hash
        }
        else {
            $excel_values_dropped += 1
        }
        Write-Progress -Id 4 -Activity "Converting to Calendar format" -Status "Processed: $excel_values_current/$excel_values_total, Dropped: $excel_values_dropped" -PercentComplete (100 / $excel_values_total * $excel_values_current)
    }
    $ProgressPreference = 'SilentlyContinue'
    foreach ($row in $calendar_values.GetEnumerator()) {
        $eventSubject = $row.Name

        if ($row.Episode){
            $row_episode = $row.Episode
            $row_name = $row.Name
            $eventDesc = "$row_name $row_episode"
        }
        else {
            $eventDesc = $row.Name
        }

        [void]$sb.AppendLine("BEGIN:VEVENT")
        [void]$sb.AppendLine("UID:" + [guid]::NewGuid())
        [void]$sb.AppendLine("CREATED:" + [datetime]::Now.ToUniversalTime().ToString($longDateFormat))
        [void]$sb.AppendLine("DTSTAMP:" + [datetime]::Now.ToUniversalTime().ToString($longDateFormat))
        [void]$sb.AppendLine("LAST-MODIFIED:" + [datetime]::Now.ToUniversalTime().ToString($longDateFormat))
        [void]$sb.AppendLine("SEQUENCE:0")
        [void]$sb.AppendLine("DTSTART:" + $row.Date)
        [void]$sb.AppendLine("DESCRIPTION:" + $eventDesc)
        [void]$sb.AppendLine("SUMMARY:" + $eventSubject)
        [void]$sb.AppendLine("LOCATION:")
        [void]$sb.AppendLine("TRANSP:TRANSPARENT")
        [void]$sb.AppendLine("END:VEVENT")
    }

    [void]$sb.AppendLine("END:VCALENDAR")

    $sb.ToString() | Out-File -Encoding "UTF8" -FilePath "$calendar_file"
}

if ($ExcelReport) {
    $excel_values | Export-Excel -Path "$excel_file" -WorksheetName "Near Expiry"
}
