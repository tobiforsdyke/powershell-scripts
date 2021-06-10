#-------------------------------------------[VARIABLES]-------------------------------------------

# SHAREPOINT URL
$siteUrl = "https://wacarts.sharepoint.com"

# SITE URL
$site = "sites/TESTTEAM9"

# LIBRARY NAME
$libraryName = "Shared Documents"

# SET TEST MODE
# Set to true to test the script (lists all the folders it would remove), if set to false script asks confirmation before deleting each folder
$whatIf = $true

# SET FORCE MODE
# Only set to true if you have fully tested it. Script WON'T ask for confirmation before deleting the file
$force = $false

#-------------------------------------------[FUNCTIONS]-------------------------------------------

Function Delete-PnPEmptyFolder([Microsoft.SharePoint.Client.Folder]$Folder)
{
    $FolderSiteRelativeURL = $Folder.ServerRelativeUrl.Substring($Web.ServerRelativeUrl.Length)

    # Process all Sub-Folders
    $SubFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderSiteRelativeURL -ItemType Folder

    Foreach($SubFolder in $SubFolders)
    {
        # Exclude "Forms" and Hidden folders
        If(($SubFolder.Name -ne "Forms") -and (-Not($SubFolder.Name.StartsWith("_"))))
        {
            # Call the function recursively
            Delete-PnPEmptyFolder -Folder $SubFolder
        }
    }

    # Get all files & Reload Sub-folders from the given Folder
    $Files = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderSiteRelativeURL -ItemType File
    $SubFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderSiteRelativeURL -ItemType Folder
 
    If ($Files.Count -eq 0 -and $SubFolders.Count -eq 0)
    {
    
    #Delete the folder
    $ParentFolder = Get-PnPProperty -ClientObject $Folder -Property ParentFolder
    $ParentFolderURL = $ParentFolder.ServerRelativeUrl.Substring($Web.ServerRelativeUrl.Length)    

    if ($whatIf -ne $true)
    {
      #Delete the folder
      Write-Host "Remove folder:" $Folder.Name "in" $ParentFolderURL -ForegroundColor Red
      Remove-PnPFolder -Name $Folder.Name -Folder $ParentFolderURL -force:$force -Recycle
    }
    else
    {
      Write-host $parentFolder
      Write-Host "Empty folder:" $Folder.Name "in" $ParentFolderURL -ForegroundColor Red
    }
    }
}

#-------------------------------------------[EXECUTION]-------------------------------------------

# Login 
$url = $siteUrl + '/' + $site
Connect-PnPOnline -Url $url -UseWebLogin

# Cleanup empty folders
$Web = Get-PnPWeb
$List = Get-PnPList -Identity $libraryName -Includes RootFolder

Delete-PnPEmptyFolder $List.RootFolder