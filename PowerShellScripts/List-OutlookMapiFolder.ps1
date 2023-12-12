param (
    $MailAddress,
    $OutputFolder
)

function Create-OutputFolder($OutputFolder) {
    if (-not (Test-Path -Path $OutputFolder)) {
        New-Item -ItemType Directory -Path $OutputFolder
    }
}
function Get-MailFolderProperties($folder) {
    # フォルダのプロパティを取得してオブジェクトに格納
    $folderProps = @{
        Name                  = $folder.Name
        AddressBookName       = $folder.AddressBookName
        CurrentView           = $folder.CurrentView
        DefaultItemType       = $folder.DefaultItemType
        DefaultMessageClass   = $folder.DefaultMessageClass
        Description           = $folder.Description
        EntryID               = $folder.EntryID
        FolderPath            = $folder.FolderPath
        InAppFolderSyncObject = $folder.InAppFolderSyncObject
        IsSharePointFolder    = $folder.IsSharePointFolder
        ItemsCount            = $folder.Items.Count
        UnReadItemCount       = $folder.UnReadItemCount
        ParentFolderName      = $folder.Parent.Name
        StoreID               = $folder.StoreID
        WebViewOn             = $folder.WebViewOn
        WebViewURL            = $folder.WebViewURL
    }

    # サブフォルダがあればそれも同様に処理
    $subFolders = $folder.Folders | ForEach-Object {
        Get-MailFolderProperties -folder $_
    }
    
    # サブフォルダがあればプロパティに追加
    if ($subFolders) {
        $folderProps['Folders'] = $subFolders
    }

    return $folderProps
}

Create-OutputFolder -OutputFolder $OutputFolder

# Outlookインスタンスを取得
$outlook = New-Object -ComObject Outlook.Application

# Outlookの名前空間を取得
$namespace = $outlook.GetNamespace("MAPI")

# 特定のメールアカウントのルートフォルダを取得
$rootFolder = $namespace.Folders.Item($MailAddress)

# メールフォルダのプロパティを含む階層構造を取得
$folderStructure = Get-MailFolderProperties -folder $rootFolder

# 階層構造をJSON形式に変換
$json = $folderStructure | ConvertTo-Json -Depth 10

# JSONを表示
#Write-Output $json
$json | Out-File "${OutputFolder}\${MailAddress}.json"
