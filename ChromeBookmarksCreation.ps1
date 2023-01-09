#Zachary McGill
#3/17/2021
#Opens the JSON bookmark file provided and adds the gibney bookmarks to the file
#Also removes the extra "Imported from IE" folders
#Now also writes to a log file everytime it is run

function CreateJSONLayout($blankFile){
    #Blank JSON Object that is needed to create the new bookmark file
    $BlankJSON = [ordered]@{
        checksum = "47a423d5cfcb5ceff4efd9cdb72c6e08"
        roots = [ordered]@{
            bookmark_bar=[ordered]@{
               children=@(
               )
               date_added = 13242404925605255
               date_modified = 13255201879641137
               guid = "00000000-0000-4000-a000-000000000002"
               id = 1
               name = "Bookmarks Bar"
               type = "Folder"
            }
            other = [ordered]@{
                children=@(
                )
                date_added = 13242404925605255
                date_modified = 0
                guid = "00000000-0000-4000-a000-000000000003"
                id = 2
                name = "Other Bookmarks"
                type = "Folder"
            }
            synced = [ordered]@{
                children=@(
                )
                date_added = 13242404925605255
                date_modified = 0
                guid = "00000000-0000-4000-a000-000000000004"
                id = 3
                name = "Mobile Bookmarks"
                type = "Folder"
            }
        }
        version = 1
    }
    #Writes the JSON to the file so the script can use it to add the bookmarks
    $blankJSON | ConvertTo-Json -Depth 3 | Out-File $localFile -Encoding utf8
    Add-Content $logFile "$(Get-Date) Write Out Blank JSON Content"
}

#Recursively finds subfolders in the bookmarks bar
function FindFolders($childNode){
    foreach($node in $childNode){
        if($node | where {$node.type -match "folder"}){
            $node.children = FindFolders($node.children)
        }
        if($node | where {$node.type -match "url"}){
            $node = NameShorting($node)
        }
    }
    return $childNode
}

#Grabs the name of the bookmark and if its length is over 256 chars cuts it to 256 chars
function NameShorting($node){
    if($node.name.Length -gt 255){
        $node.name = $node.name.subString(0, [System.Math]::Min(255, $node.name.Length))
        Add-Content $logFile "$(Get-Date) Shortened Name: $node" 
        return $node
    }
    return $node
}

#Create the bookmarks as JSON objects to append if they are needed
$Zendesk = @(
    @{
         date_added = '13242404925605255'
         name = 'Gibney Support'
         show_icon = $false
         source = 'user_add'
         type = 'url'
         url = '###REDACTED###'
    }
)

$Intranet = @(
    @{
         date_added = '13242404925605255'
         name = 'Gibney Intranet'
         show_icon = $false
         source = 'user_add'
         type = 'url'
         url = '###REDACTED###'
    }
)

$ProofPoint = @(
    @{
         date_added = '13242404925605255'
         name = 'ProofPoint Essentials Spam Filter'
         show_icon = $false
         source = 'user_add'
         type = 'url'
         url = '###REDACTED###'
    }
)

$Zoom = @(
	@{
		date_added = '13242404925605255'
		name = 'Zoom'
		show_icon = $false
		source = 'user_add'
		type = 'url'
		url = '###REDACTED###'
	}
)

$Tracker = @(
    @{
         date_added = '13242404925605255'
         name = 'Tracker 8'
         show_icon = $false
         source = 'user_add'
         type = 'url'
         url = '###REDACTED###'
    }
)

$Paycom = @(
	@{
		date_added = '13242404925605255'
		name = 'Paycom'
		show_icon = $false
		source = 'user_add'
		type = 'url'
		url = '###REDACTED###'
	}
)

$Robin = @(
	@{
		date_added = '13242404925605255'
		name = 'Robin'
		show_icon = $false
		source = 'user_add'
		type = 'url'
		url = '###REDACTED###'
	}
)

#Enables logging
$logging = $true

#Set the log folder path
$logFolder = 'C:\BK Logs'

$userFolder = 'H:\ChromeBookmarks'

#check if logging is enabled
if($logging = $true){
	#Test if the folder exists
	if(!(test-path $logFolder)){
		#Create new folder
		New-Item -ItemType Directory -Force -Path $logFolder
	}
	#Grabs the username to use in the file name
    $User = $env:USERNAME
    #Create a new file in the logging folder using the date and time as the name of the file
	$FileName = $User + (Get-Date).tostring(" dd-MM-yyyy-hh-mm-ss") + ".log"
    #Creates the file using fileName
	New-Item -ItemType File -Path $logFolder -Name ($FileName)
    #Stores the file path to use later
	$logFile = $logFolder + "\" + $FileName
	Add-Content $logFile "$(Get-Date) Created Log File"
}

#Test if the folder exists
if(!(test-path $userFolder)){
	#Create new folder
	New-Item -ItemType Directory -Force -Path $userFolder
    Add-Content $logFile "$(Get-Date) Added H drive folder"
}

#File Chrome uses to open bookmarks
#The ~ is the same thing as C:\Users\%Username%\
$localFile = '~\AppData\Local\Google\Chrome\User Data\Default\Bookmarks'

#Check to see the bookmarks exist
if(!(test-path $localFile)){
    #Create new bookmarks file
    New-Item -ItemType File -Path $localFile
    Add-Content $logFile "$(Get-Date) Created new bookmarks file"
    #Add blank JSON object so we can write default bookmarks to file
    CreateJsonLayout($localFile)
}

#Check to see if the file is greater than 0KB
$counter =0
#Get the size of the file
$beforeFileSize = (Get-Item $localFile).Length
while(($beforeFileSize -eq 0) -AND $counter -lt 5) {
    $counter++
    if($counter -eq 5){
        #We've waited 5 seconds and the file is still empty, create a blank file
        Add-Content $logFile "$(Get-Date) Zero byte file, creating Empty file"
        CreateJSONLayout($localFile)
    }
    Add-Content $logFile "$(Get-Date) Sleep $counter second(s)"
    Start-Sleep -Seconds 1
}
Add-Content $logFile "$(Get-Date) Opened Bookmarks File"

#Read in the data from file and convert it to PSObject instead of JSON
$jsonContent = Get-Content -Raw -Path $localFile | ConvertFrom-Json
Add-Content $logFile "$(Get-Date) Converted from JSON: $($jsonContent.roots.bookmark_bar.children.PSobject.Properties.Value.Count)"

#Variables to store if the bookmarks should be added
$supportCheck = $false
$intranetCheck = $false
$proofpointCheck = $false
$zoomCheck = $false
$trackerCheck = $false
$paycomCheck = $false
$robinCheck = $false

#Initialize object to add bookmarks
$testBookmark = @()

#Check if the URL is already present in the bookmarks bar
foreach($bookmark in $jsonContent.roots.bookmark_bar.children){
    Add-Content $logFile "$(Get-Date) Checking bookmark; Count: $($jsonContent.roots.bookmark_bar.children.PSobject.Properties.Value.Count)"
    if($bookmark.url -match "###REDACTED###"){
        if($supportCheck -eq $false){
            $supportCheck = $true
            $testBookmark += $bookmark
            Add-Content $logFile "$(Get-Date) --Support found! Count: $($testBookmark.PSobject.Properties.Value.Count)"
        }else{
            Add-Content $logFile "$(Get-Date) --Support found again! Not adding to Bookmark file. Count: $($testBookmark.PSobject.Properties.Value.Count)"
        }
    }elseif($bookmark.url -match "###REDACTED###"){
        if($intranetCheck -eq $false){
            $intranetCheck = $true
            $testBookmark += $bookmark
            Add-Content $logFile "$(Get-Date) --Intranet found! Count: $($testBookmark.PSobject.Properties.Value.Count)"
        }else{
            #log - found another Tracker, not adding
            Add-Content $logFile "$(Get-Date) --Intranet found again! Not adding to Bookmark file. Count: $($testBookmark.PSobject.Properties.Value.Count)"
        }
    }elseif($bookmark.url -match "###REDACTED###"){
        if($proofpointCheck -eq $false){
            $proofpointCheck = $true
            $testBookmark += $bookmark
            Add-Content $logFile "$(Get-Date) --Proofpoint found! Count: $($testBookmark.PSobject.Properties.Value.Count)"
        }else{
            #log - found another Tracker, not adding
            Add-Content $logFile "$(Get-Date) --Proofpoint found again! Not adding to Bookmark file. Count: $($testBookmark.PSobject.Properties.Value.Count)"
        }
    }elseif($bookmark.url -match "###REDACTED###"){
        if($zoomCheck -eq $false){
            $zoomCheck = $true
            $testBookmark += $bookmark
            Add-Content $logFile "$(Get-Date) --Zoom found! Count: $($testBookmark.PSobject.Properties.Value.Count)"
        }else{
            #log - found another Tracker, not adding
            Add-Content $logFile "$(Get-Date) --Zoom found again! Not adding to Bookmark file. Count: $($testBookmark.PSobject.Properties.Value.Count)"
        }
    }elseif($bookmark.url -match "###REDACTED###"){
        if($trackerCheck -eq $false){
            $trackerCheck = $true
            $testBookmark += $bookmark
            Add-Content $logFile "$(Get-Date) --Tracker found! Count: $($testBookmark.PSobject.Properties.Value.Count)"
        }else{
            #log - found another Tracker, not adding
            Add-Content $logFile "$(Get-Date) --Tracker found again! Not adding to Bookmark file. Count: $($testBookmark.PSobject.Properties.Value.Count)"
        }
    }elseif($bookmark.url -match "###REDACTED###"){
        if($paycomCheck -eq $false){
            $paycomCheck = $true
            $testBookmark += $bookmark
            Add-Content $logFile "$(Get-Date) --Paycom found! Count: $($testBookmark.PSobject.Properties.Value.Count)"
        }else{
            #log - found another Tracker, not adding
            Add-Content $logFile "$(Get-Date) --Paycom found again! Not adding to Bookmark file. Count: $($testBookmark.PSobject.Properties.Value.Count)"
        }
    }elseif($bookmark.url -match "###REDACTED###"){
        if($robinCheck -eq $false){
            $robinCheck = $true
            $testBookmark += $bookmark
            Add-Content $logFile "$(Get-Date) --Robin found! Count: $($testBookmark.PSobject.Properties.Value.Count)"
        }else{
            #log - found another Tracker, not adding
            Add-Content $logFile "$(Get-Date) --Robin found again! Not adding to Bookmark file. Count: $($testBookmark.PSobject.Properties.Value.Count)"
        }
    }else{
        if($bookmark.name -match "Imported From IE \("){
            Add-Content $logFile "$(Get-Date) --Removed Imported Folder Count: $($jsonContent.roots.bookmark_bar.children.PSobject.Properties.Value.Count)"
        }elseif($bookmark.url -match "###REDACTED###"){
			Add-Content $logFile "$(Get-Date) --Removed Paychex Count: $($jsonContent.roots.bookmark_bar.children.PSobject.Properties.Value.Count)"
		}else{
            #Add Bookmark to new bookmarks object
            $testBookmark += $bookmark
        }
    }
}

#If the URLs are not found; add them to the bookmarks file
if($supportCheck -eq $false){   
    $testBookmark += $Zendesk
    Add-Content $logFile "$(Get-Date) --Adding Zendesk to the Bookmarks! Count: $($jsonContent.roots.bookmark_bar.children.PSobject.Properties.Value.Count)"
}
if($intranetCheck -eq $false){
    $testBookmark += $Intranet
    Add-Content $logFile "$(Get-Date) --Adding Intranet to the Bookmarks! Count: $($jsonContent.roots.bookmark_bar.children.PSobject.Properties.Value.Count)"
}
if($proofpointCheck -eq $false){
    $testBookmark += $ProofPoint
    Add-Content $logFile "$(Get-Date) --Adding Proofpoint to the Bookmarks! Count: $($jsonContent.roots.bookmark_bar.children.PSobject.Properties.Value.Count)"
}
if($zoomCheck -eq $false){
	$testBookmark += $Zoom
    Add-Content $logFile "$(Get-Date) --Adding Zoom to the Bookmarks! Count: $($jsonContent.roots.bookmark_bar.children.PSobject.Properties.Value.Count)"
}
if($trackerCheck -eq $false){
	$testBookmark += $Tracker
    Add-Content $logFile "$(Get-Date) --Adding Tracker to the Bookmarks! Count: $($jsonContent.roots.bookmark_bar.children.PSobject.Properties.Value.Count)"
}
if($paycomCheck -eq $false){
	$testBookmark += $Paycom
    Add-Content $logFile "$(Get-Date) --Adding Paycom to the Bookmarks! Count: $($jsonContent.roots.bookmark_bar.children.PSobject.Properties.Value.Count)"
}
if($robinCheck -eq $false){
	$testBookmark += $Robin
    Add-Content $logFile "$(Get-Date) --Adding Robin to the Bookmarks! Count: $($jsonContent.roots.bookmark_bar.children.PSobject.Properties.Value.Count)"
}
#Set the JSON Object to the bookmarks without "Imported from IE" folders
$jsonContent.roots.bookmark_bar.children = $testBookmark
Add-Content $logFile "$(Get-Date) Wrote the bookmarks back to the object: $($jsonContent.roots.bookmark_bar.children.PSobject.Properties.Value.Count)"

#Calls the find folders to do the name shortening
foreach($child in $jsonContent.roots.bookmark_bar.children){
    FindFolders($child) | Out-Null
}

#Converts to JSON and writes object back out to the file
$jsonContent | ConvertTo-Json -Depth 50 | Out-File $localFile -Encoding utf8
Add-Content $logFile "$(Get-Date) Converted back to JSON; Count: $($jsonContent.roots.bookmark_bar.children.PSobject.Properties.Value.Count)"

#Logs the file size after running the script
$afterFileSize = (Get-Item $localFile).Length
Add-Content $logFile "$(Get-Date) File Size Before running: $beforeFileSize bytes, After runnning: $afterFileSize bytes; Count: $($jsonContent.roots.bookmark_bar.children.PSobject.Properties.Value.Count)"
Add-Content $logFile "$(Get-Date) Wrote the file back out`n ====Finished===="