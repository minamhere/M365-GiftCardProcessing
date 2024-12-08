# Parameters
$ParentFolderName = "Gift Cards"
$SubFolderName = "Unprocessed"

# Script
Write-Output "Starting script..."

Begin {
    # Function to handle Costco BHN emails
    function Handle-CostcoBHNEmail {
        param (
            [string]$Subject,
            [string]$Body
        )

        Write-Output "Handling Costco BHN email..."
        Write-Output "Subject: $Subject"
        # Add logic here to process the email content
    }

    # Function to handle other types of emails (placeholder)
    function Handle-OtherEmails {
        param (
            [string]$Sender,
            [string]$Subject,
            [string]$Body
        )

        Write-Output "Handling other email..."
        Write-Output "Sender: $Sender"
        Write-Output "Subject: $Subject"
        # Add logic here to process other email types
    }
}

Process {
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
        $Subject = $Message.Subject
        $Body = $Message.Body.Content

        if ($Sender -eq "donotreply.giftcards.costco@bhnetwork.com") {
            # Call the Costco BHN handler function
            Handle-CostcoBHNEmail -Subject $Subject -Body $Body
        } else {
            Write-Output "Unknown email sender: $Sender. Exiting for now."
            break
        }
    }

    # Disconnect from Microsoft Graph
    Disconnect-MgGraph
    Write-Output "Disconnected from Microsoft Graph."
}

End {
    Write-Output "Script completed."
}
