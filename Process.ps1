# Parameters
$ParentFolderName = "Gift Cards"
$SubFolderName = "Unprocessed"

# Function to handle Costco BHN emails
function Handle-CostcoBHNEmail {
    param (
        [string]$Body
    )

    # Regex to extract the embedded link
    $Regex = 'https:%2F%2Fegift\.activationspot\.com%2F[^\s"]*'
    $Match = [regex]::Match($Body, $Regex)

    if ($Match.Success) {
        $EmbeddedLink = $Match.Value

        # Decode %2F to /, %26 to &, and %3F to ?
        $CleanedLink = $EmbeddedLink -replace '%2F', '/' -replace '%26', '&' -replace '%3F', '?'

        # Remove the trailing slash and everything after the tid parameter
        $CleanedLink = $CleanedLink -replace '(&tid=[^\/]+)\/.*$', '$1'

        # Print the cleaned link
        Write-Output $CleanedLink
    } else {
        Write-Output "No valid links found in this email."
    }
}





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

# Step 6: Fetch all messages in the subfolder
Write-Output "Fetching all messages from subfolder '$SubFolderName'..."
$Messages = Get-MgUserMailFolderMessage -MailFolderId $SubFolderId -UserId $Account -All

Write-Output "Retrieved $(($Messages | Measure-Object).Count) messages."

# Step 7: Process each email
foreach ($Message in $Messages) {
    $Sender = $Message.From.EmailAddress.Address
    $Body = $Message.Body.Content

    if ($Sender -eq "donotreply.giftcards.costco@bhnetwork.com") {
        # Call the Costco BHN handler function
        Handle-CostcoBHNEmail -Body $Body
    } else {
        Write-Output "Unknown email sender: $Sender. Exiting for now."
        break
    }
}

# Disconnect from Microsoft Graph
Disconnect-MgGraph
Write-Output "Disconnected from Microsoft Graph."
