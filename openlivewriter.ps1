#download zip file for syntax highlighting
$url = "https://richhewlett.blob.core.windows.net/blogdownloads/SyntaxHighlight_WordPressCom_OLWPlugIn_V2.0.0.zip" 
$path = (Get-Item -Path ".\" -Verbose).FullName + "\SyntaxHighlight_WordPressCom_OLWPlugIn_V2.0.0.zip"

Write-Host "Downloading [$url]`nSaving at [$path]" 
$client = new-object System.Net.WebClient 
$client.DownloadFile($url, $path ) 
      
#create destination directory
$LocalAppDir = [Environment]::GetFolderPath("LocalApplicationData")
$OpenliveWriterDir = $LocalAppDir + "\OpenLiveWriter"

#just in case there are multiple versions installed, or there's an upgrade, enumerate the application binary folders
$appDirectories = Get-ChildItem $OpenliveWriterDir | Where-Object {$_ -like "app-*"};

foreach ($appDirectory in $appDirectories) {

    #Create destination directory
    $targetDir = $appDirectory.FullName + "\Plugins"
    if(!(Test-Path -Path $targetDir )){
        Write-Host "Creating path $targetDir"
        New-Item -ItemType directory -Path $targetDir
    }


    #decompress file into target directory
    Write-Host "Expanding $path to $targetDir"
    Expand-Archive $path -DestinationPath $targetDir

    #update config file
    $configFileName = $appDirectory.FullName + "\OpenLiveWriter.exe.config"
    [xml] $xml = gc $configFileName

    #create the new element
    $newitem = $xml.CreateElement("loadFromRemoteSources")
    $newitem.SetAttribute("enabled","true")

    #if it doesn't already exist in the file, add it and save
    if (!$xml.SelectSingleNode("//loadFromRemoteSources")) {
        $xml.DocumentElement.LastChild.AppendChild($newitem)
        $xml.Save($configFileName)
    }    
}


