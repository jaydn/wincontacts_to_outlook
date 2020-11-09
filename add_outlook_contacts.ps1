Get-ChildItem .\outputs\*.csv | ForEach-Object {
    $outlook = new-object -com Outlook.Application
    $contacts = $outlook.Session.GetDefaultFolder(10)
    $session = $outlook.Session
    $session.Logon("Outlook")
    $namespace = $outlook.GetNamespace("MAPI")
    $toAddRecipients = @()

    Import-Csv -Path $_ -Header Email, Name| Foreach-Object {
            $recipient = $namespace.CreateRecipient("$($_.Name) <$($_.Email)>")
            $toAddRecipients += $recipient
    }

    $toAddRecipients | ForEach-Object { $_.Resolve(); Write-Host $_ }
    $DL = $contacts.Items.Add("IPM.DistList")
    $DL.DLName = [System.IO.Path]::GetFileNameWithoutExtension($_)
    $toAddRecipients | ForEach-Object { $DL.AddMember($_) }
    $DL.Save()
}
