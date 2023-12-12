[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
$flowLayoutPanel = New-Object Windows.Forms.FlowLayoutPanel
$flowLayoutPanel.AutoSize = $true
$flowLayoutPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
$form = New-Object Windows.Forms.Form
$form.Text = "What are you going to do .."
$form.Controls.Add($flowLayoutPanel)
# Display the form as a dialog

function Add-Button {
    param (
        $Function,
        $Text
    )
    $button = New-Object Windows.Forms.Button
    $button.AutoSize = $true
    $button.Text = $Text
    $button.Add_Click($Function)
    $flowLayoutPanel.Controls.Add($button)
}

$EmailAddress = "test@example.com"
Add-Button { 
    #[System.Windows.Forms.MessageBox]::Show("Menu1")
    .\PowerShellScripts\List-ExchangeFolder.ps1 `
        -OutputFolder ExchangeFolder `
        -UserPrincipalName $EmailAddress
} "Outlook ${EmailAddress}"

$EmailAddress = "foo@example.com"
Add-Button  {
    #[System.Windows.Forms.MessageBox]::Show("Menu1")
    .\PowerShellScripts\List-OutlookMapiFolder.ps1 `
        -OutputFolder OutlookMapiFolder `
        -MailAddress $EmailAddress
} "Outlook ${EmailAddress}"

$form.ShowDialog()
$form.Dispose()
