# Parameters
$ClientId = "YOUR_APP_CLIENT_ID"   # Replace with your Azure App's Client ID
$TenantId = "YOUR_TENANT_ID"       # Replace with your Azure App's Tenant ID

$ParentFolderName = "Gift Cards"
$SubFolderName = "Unprocessed"

# Get an access token interactively
$TokenResponse = Get-MsalToken -ClientId $ClientId -TenantId $TenantId -RedirectUri "https://login.microsoftonline.com/common/oauth2/nativeclient" -Scopes "Mail.Read"
$AccessToken = $TokenResponse.AccessToken

# Define API endpoints
$MailFoldersEndpoint = "https://graph.microsoft.com/v1.0/me/mailFolders"
$MessagesEndpointTemplate = "https://graph.microsoft.com/v1.0/me/mailFolders/{folder-id}/messages"

# Function to find a folder by name within a parent folder
function Get-ChildFolderId {
    param (
        [string]$ParentFolderId,
        [string]$ChildFolderName
    )

    $ChildFoldersEndpoint = "https://graph.microsoft.com/v1.0/me/mailFolders/$ParentFolderId/childFolders"
    $ChildFolders = Invoke-RestMethod -Uri $ChildFoldersEndpoint -Headers @{Authorization = "Bearer $AccessToken"} -Method Get

    $ChildFolder = $ChildFolders.value | Where-Object { $_.displayName -eq $ChildFolderName }

    if ($ChildFolder) {
        return $ChildFolder.id
    } else {
        return $null
    }
}

# Step 1: Get the parent folder ID
Write-Output "Fetching mail folders..."
$MailFolders = Invoke-RestMethod -Uri $MailFoldersEndpoint -Headers @{Authorization = "Bearer $AccessToken"} -Method Get

$ParentFolder = $MailFolders.value | Where-Object { $_.displayName -eq $ParentFolderName }

if (-not $ParentFolder) {
    Write-Output "Parent folder '$ParentFolderName' not found. Exiting script."
    exit
}

$ParentFolderId = $ParentFolder.id
Write-Output "Found parent folder '$ParentFolderName' with ID: $ParentFolderId"

# Step 2: Get the subfolder ID
$SubFolderId = Get-ChildFolderId -ParentFolderId $ParentFolderId -ChildFolderName $SubFolderName

if (-not $SubFolderId) {
    Write-Output "Subfolder '$SubFolderName' under '$ParentFolderName' not found. Exiting script."
    exit
}

Write-Output "Found subfolder '$SubFolderName' with ID: $SubFolderId"

# Step 3: Fetch messages in the subfolder
$MessagesEndpoint = $MessagesEndpointTemplate -replace "{folder-id}", $SubFolderId

$Messages = Invoke-RestMethod -Uri $MessagesEndpoint -Headers @{Authorization = "Bearer $AccessToken"} -Method Get

# Process each email
foreach ($Message in $Messages.value) {
    $Subject = $Message.subject
    $Sender = $Message.sender.emailAddress.address
    $Body = $Message.body.content

    Write-Output "Processing email..."
    Write-Output "Subject: $Subject"
    Write-Output "Sender: $Sender"
    Write-Output "Body: $Body"
    Write-Output "========================="
}
