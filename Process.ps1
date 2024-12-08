# Parameters
$ParentFolderName = "Gift Cards"
$SubFolderName = "Unprocessed"

# Step 1: Connect to Microsoft Graph
Write-Output "Authenticating to Microsoft Graph..."
Connect-MgGraph -Scopes "Mail.Read"

# Step 2: Retrieve the authenticated user's account (email address)
Write-Output "Fetching authenticated user's account..."
$Account = (Get-MgContext).Account

if (-not $Account) {
    Write-Output "Failed to retrieve the authenticated user's account. Exiting script."
    Disconnect-MgGraph
    exit
}

Write-Output "Authenticated user's account: $Account"

# Step 3: Get all mail folders for the authenticated user
Write-Output "Fetching mail folders..."
$MailFolders = Get-MgUserMailFolder -UserId $Account

# Step 4: Find the parent folder
$ParentFolder = $MailFolders | Where-Object { $_.DisplayName -eq $ParentFolderName }

if (-not $ParentFolder) {
    Write-Output "Parent folder '$ParentFolderName' not found. Exiting script."
    Disconnect-MgGraph
    exit
}

$ParentFolderId = $ParentFolder.Id
Write-Output "Found parent folder '$ParentFolderName' with ID: $ParentFolderId"

# Step 5: Find the subfolder within the parent folder
$ChildFolders = Get-MgUserMailFolderChildFolder -MailFolderId $ParentFolderId -UserId $Account
$SubFolder = $ChildFolders | Where-Object { $_.DisplayName -eq $SubFolderName }

if (-not $SubFolder) {
    Write-Output "Subfolder '$SubFolderName' under '$ParentFolderName' not found. Exiting script."
    Disconnect-MgGraph
    exit
}

$SubFolderId = $SubFolder.Id
Write-Output "Found subfolder '$SubFolderName' with ID: $SubFolderId"

# Step 6: Fetch messages in the subfolder
Write-Output "Fetching messages from subfolder '$SubFolderName'..."
$Messages = Get-MgUserMailFolderMessage -MailFolderId $SubFolderId -UserId $Account -All

# Step 7: Process each email
foreach ($Message in $Messages) {
    $Subject = $Message.Subject
    $Sender = $Message.From.EmailAddress.Address
    $Body = $Message.Body.Content

    Write-Output "Processing email..."
    Write-Output "Subject: $Subject"
    Write-Output "Sender: $Sender"
    Write-Output "Body: $Body"
    Write-Output "========================="
}

# Disconnect from Microsoft Graph
Disconnect-MgGraph
Write-Output "Disconnected from Microsoft Graph."
