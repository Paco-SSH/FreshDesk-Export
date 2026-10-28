
# ============================================
# Freshdesk - Complete Excel Export
# Tickets + Notes
# Will need to enter your FreshDesk URL below
# Script will ask you for API key manually
# Will need to update date if using since before 2020-01-01
# Will stop and wait 60 seconds if API limit is hit
# ============================================

$ErrorActionPreference = "Stop"

# Configuration
$Domain = "URL_HERE"
$Output = "$env:USERPROFILE\Downloads\Freshdesk_Export.xlsx"
$UpdatedSince = "2010-10-13T00:00:00Z"

# Install ImportExcel if required
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module ImportExcel -Scope CurrentUser -Force
}

Import-Module ImportExcel

# Authentication
$ApiKey = (Read-Host "Enter Freshdesk API key").Trim()

$Auth = [Convert]::ToBase64String(
    [Text.Encoding]::ASCII.GetBytes("${ApiKey}:X")
)

$Headers = @{
    Authorization = "Basic $Auth"
    Accept = "application/json"
}

# API helper with retry handling
function Get-FreshdeskData {
    param([string]$Url)

    for ($Attempt = 1; $Attempt -le 5; $Attempt++) {
        try {
            return Invoke-RestMethod `
                -Uri $Url `
                -Headers $Headers `
                -Method Get `
                -ErrorAction Stop
        }
        catch {
            $StatusCode = [int]$_.Exception.Response.StatusCode

            if ($StatusCode -eq 429 -and $Attempt -lt 5) {
                Write-Warning "Rate limited. Retrying in 60 seconds..."
                Start-Sleep -Seconds 60
                continue
            }

            throw
        }
    }
}

# Collections
$Tickets = [System.Collections.Generic.List[object]]::new()
$Notes = [System.Collections.Generic.List[object]]::new()



Write-Host "Retrieving Freshdesk tickets..." -ForegroundColor Cyan

$Page = 1
$PerPage = 100

do {
    $Url = "https://$Domain/api/v2/tickets" +
           "?updated_since=$UpdatedSince" +
           "&include=description" +
           "&per_page=$PerPage&page=$Page"

    $Result = Get-FreshdeskData -Url $Url
    $PageTickets = @($Result)

    foreach ($Ticket in $PageTickets) {

        $Status = switch ($Ticket.status) {
            2 { "Open" }
            3 { "Pending" }
            4 { "Resolved" }
            5 { "Closed" }
            default { "$($Ticket.status)" }
        }

        $Priority = switch ($Ticket.priority) {
            1 { "Low" }
            2 { "Medium" }
            3 { "High" }
            4 { "Urgent" }
            default { "$($Ticket.priority)" }
        }

        $Tickets.Add([PSCustomObject]@{
            ID          = $Ticket.id
            Subject     = $Ticket.subject
            Description = $Ticket.description_text
            Status      = $Status
            Priority    = $Priority
            RequesterID = $Ticket.requester_id
            AgentID     = $Ticket.responder_id
            GroupID     = $Ticket.group_id
            Type        = $Ticket.type
            Created     = $Ticket.created_at
            Updated     = $Ticket.updated_at
            DueBy       = $Ticket.due_by
        })
    }

    Write-Host "Page $Page - Total tickets: $($Tickets.Count)"

    $Page++

} while ($PageTickets.Count -eq $PerPage -and $Page -le 300)

if ($Page -gt 300 -and $PageTickets.Count -eq $PerPage) {
    Write-Warning "Reached Freshdesk's 300-page limit. Additional tickets may exist."
}

Write-Host "Retrieved $($Tickets.Count) tickets." -ForegroundColor Green



Write-Host "Retrieving ticket notes..." -ForegroundColor Cyan

$Counter = 0

foreach ($Ticket in $Tickets) {

    $Counter++
    $TicketID = $Ticket.ID

    Write-Progress `
        -Activity "Retrieving Notes" `
        -Status "Ticket $Counter of $($Tickets.Count)" `
        -PercentComplete (($Counter / $Tickets.Count) * 100)

    $NotePage = 1

    do {
        $Url = "https://$Domain/api/v2/tickets/$TicketID/conversations" +
               "?per_page=100&page=$NotePage"

        $Result = Get-FreshdeskData -Url $Url
        $Conversations = @($Result)

        foreach ($Conversation in $Conversations) {

            # Source 2 = Freshdesk note
            if ($Conversation.source -eq 2) {

                # Excel cell text limit
                $NoteText = [string]$Conversation.body_text

                if ($NoteText.Length -gt 32767) {
                    $NoteText = $NoteText.Substring(0, 32767)
                }

                $Notes.Add([PSCustomObject]@{
                    TicketID = $TicketID
                    Subject  = $Ticket.Subject
                    NoteID   = $Conversation.id
                    Note     = $NoteText
                    Private  = $Conversation.private
                    AuthorID = $Conversation.user_id
                    Created  = $Conversation.created_at
                    Updated  = $Conversation.updated_at
                })
            }
        }

        $NotePage++

    } while ($Conversations.Count -eq 100)

    # Reduce API rate-limit pressure
    Start-Sleep -Milliseconds 250
}

Write-Progress -Activity "Retrieving Notes" -Completed

Write-Host "Retrieved $($Notes.Count) notes." -ForegroundColor Green



Write-Host "Creating Excel workbook..." -ForegroundColor Cyan

if (Test-Path $Output) {
    Remove-Item $Output -Force
}

# Sheet 1: Tickets
if ($Tickets.Count -gt 0) {
    $Tickets | Export-Excel `
        -Path $Output `
        -WorksheetName "Tickets" `
        -AutoSize `
        -AutoFilter `
        -FreezeTopRow `
        -BoldTopRow
}
else {
    throw "No tickets were returned. No workbook created."
}

# Sheet 2: Notes
if ($Notes.Count -gt 0) {
    $Notes | Export-Excel `
        -Path $Output `
        -WorksheetName "Notes" `
        -AutoSize `
        -AutoFilter `
        -FreezeTopRow `
        -BoldTopRow
}
else {
    # Create empty Notes sheet
    $Excel = Open-ExcelPackage -Path $Output
    Add-Worksheet -ExcelPackage $Excel -WorksheetName "Notes" | Out-Null
    Close-ExcelPackage $Excel
}

# ============================================
# COMPLETE
# ============================================

Write-Host ""
Write-Host "=====================================" -ForegroundColor Green
Write-Host "Freshdesk Export Completed"
Write-Host "=====================================" -ForegroundColor Green
Write-Host "Tickets Exported: $($Tickets.Count)"
Write-Host "Notes Exported:   $($Notes.Count)"
Write-Host "Output: $Output"
Write-Host "=====================================" -ForegroundColor Green
