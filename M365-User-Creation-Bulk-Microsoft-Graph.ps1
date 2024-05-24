<#
.SYNOPSIS
    Script to create users in Microsoft 365 using Microsoft Graph API.

.DESCRIPTION
    This PowerShell script automates the creation of users in Microsoft 365 based on a CSV file input. 
    It uses the Microsoft Graph API to create users with specified properties and includes robust 
    error handling, logging, and progress tracking.

.PARAMETER CSVFilePath
    The path to the CSV file containing user details. The CSV should have the following columns:
    DisplayName, UserPrincipalName, Password, First Name, Last Name, Job title, Department, 
    Usage location, State, Country, Office Location, City, Postal Code.

.NOTES
    Author: Sivakumar Margabandhu from m365simplified.com
    Date: 2024-05-23
    Version: 1.0
    Requires: Microsoft.Graph PowerShell module

.EXAMPLE
    # Run the script with a specified CSV file
    .\Create-M365Users.ps1 -CSVFilePath "path-to-your-csv-file.csv"

#>

param (
    [Parameter(Mandatory=$true)]
    [string]$CSVFilePath
)

# Import necessary module
Import-Module Microsoft.Graph

# Define log file path
$logFilePath = "user_creation_log.txt"

# Function to log messages
function Log-Message {
    param (
        [string]$message,
        [string]$type = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$type] $message"
    Write-Output $logMessage
    Add-Content -Path $logFilePath -Value $logMessage
}

# Connect to Microsoft Graph
try {
    Connect-MgGraph -Scopes "User.ReadWrite.All" -ErrorAction Stop
    Log-Message "Successfully connected to Microsoft Graph."
} catch {
    Log-Message "Failed to connect to Microsoft Graph. Error: $_" "ERROR"
    throw
}

# Import users from CSV
try {
    $users = Import-Csv -Path $CSVFilePath
    Log-Message "Successfully imported users from CSV file."
} catch {
    Log-Message "Failed to import users from CSV file. Error: $_" "ERROR"
    throw
}

$totalUsers = $users.Count
$createdUsers = 0

# Display progress bar setup
$progressBar = @{
    Activity = "Creating Users in Microsoft 365"
    Status = "Initializing"
    PercentComplete = 0
    CurrentOperation = ""
}

# Process each user
foreach ($user in $users) {
    try {
        $progressBar.CurrentOperation = "Creating user: $($user.DisplayName)"
        $progressBar.PercentComplete = [math]::Round(($createdUsers / $totalUsers) * 100, 2)
        Write-Progress @progressBar

        # Create password profile
        $passwordProfile = @{
            Password = $user.Password
            ForceChangePasswordNextSignIn = $true
        }

        # Create user parameters
        $userParams = @{
            AccountEnabled    = $true
            DisplayName       = $user.DisplayName
            MailNickname      = $user.UserPrincipalName.Split('@')[0]
            UserPrincipalName = $user.UserPrincipalName
            PasswordProfile   = $passwordProfile
            GivenName         = $user."First Name"
            Surname           = $user."Last Name"
            JobTitle          = $user."Job title"
            Department        = $user.Department
            UsageLocation     = $user."Usage location"
            OfficeLocation    = $user."Office Location"
            City              = $user.City
            State             = $user.State
            Country           = $user.Country
            PostalCode        = $user."Postal Code"
        }

        # Create user in Microsoft 365 using named parameters
        New-MgUser @userParams
        $createdUsers++

        # Log success
        Log-Message "Successfully created user: $($user.DisplayName) ($($user.UserPrincipalName))"

    } catch {
        # Log error
        Log-Message "Failed to create user: $($user.DisplayName) ($($user.UserPrincipalName)). Error: $_" "ERROR"
    }

    # Display progress
    Write-Output "Created $createdUsers out of $totalUsers users"

    # Pause for 3 seconds after every 20 users
    if ($createdUsers % 20 -eq 0) {
        Start-Sleep -Seconds 3
    }
}

# Final progress update
$progressBar.PercentComplete = 100
$progressBar.CurrentOperation = "Completed creating all users."
Write-Progress @progressBar -Completed

# Display completion message
Log-Message "Completed creating $createdUsers users out of $totalUsers users."
Write-Output "Completed creating $createdUsers users out of $totalUsers users."
