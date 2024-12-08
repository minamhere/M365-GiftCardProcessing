# Parameters
$ParentFolderName = "Gift Cards"
$SubFolderName = "Unprocessed"

# Step 1: Connect to Microsoft Graph
Write-Output "Authenticating to Microsoft Graph..."
Connect-MgGraph -Scopes "Mail.Read"

# Step 2: Get all mail folders
Write-Output "Fetching mail folders..."
$MailFolders = Get-MgUserMailFolder -UserId "me"

# Step 3: Find the parent folder
$ParentFolder = $MailFolders | Where-Object { $_.DisplayName -eq $ParentFolderName }

if (-not $ParentFolder) {
    Write-Output "Parent folder '$ParentFolderName' not found. Exiting script."
    Disconnect-MgGraph
    exit
}

$ParentFolderId = $ParentFolder.Id
Write-Output "Found parent folder '$ParentFolderName' with ID: $ParentFolderId"

# Step 4: Find the subfolder within the parent folder
$ChildFolders = Get-MgUserMailFolderChildFolder -UserId "me" -MailFolderId $ParentFolderId
$SubFolder = $ChildFolders | Where-Object { $_.DisplayName -eq $SubFolderName }

if (-not $SubFolder) {
    Write-Output "Subfolder '$SubFolderName' under '$ParentFolderName' not found. Exiting script."
    Disconnect-MgGraph
    exit
}

$SubFolderId = $SubFolder.Id
Write-Output "Found subfolder '$SubFolderName' with ID: $SubFolderId"

# Step 5: Fetch messages in the subfolder
Write-Output "Fetching messages from subfolder '$SubFolderName'..."
$Messages = Get-MgUserMailFolderMessage -UserId "me" -MailFolderId $SubFolderId

# Step 6: Process each email
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
