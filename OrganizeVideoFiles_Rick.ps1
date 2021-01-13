# ==============================================================================================
# 
# Microsoft PowerShell Source File 
# 
# This script will organize photo and video files by renaming the file based on the date the
# file was created and moving them into folders based on the year and month. It will also append
# a random number to the end of the file name just to avoid name collisions. The script will
# look in the SourceRootPath (recursing through all subdirectories) for any files matching
# the extensions in FileTypesToOrganize. It will rename the files and move them to folders under
# DestinationRootPath, e.g. DestinationRootPath\2011\02_February\2011-02-09_21-41-47_680.jpg
# 
# JPG files contain EXIF data which has a DateTaken value. Other media files have a MediaCreated
# date. 
#
# The code for extracting the EXIF DateTaken is based on a script by Kim Oppalfens:
# http://blogcastrepository.com/blogs/kim_oppalfenss_systems_management_ideas/archive/2007/12/02/organize-your-digital-photos-into-folders-using-powershell-and-exif-data.aspx
# ============================================================================================== 

#[reflection.assembly]::loadfile( "C:\Windows\Microsoft.NET\Framework\v2.0.50727\System.Drawing.dll") 

$SourceRootPath = "M:\PhotosNonClasse"
$DestinationRootPath = "M:\videos"
$FileTypesToOrganize = @("*.mov","*.avi","*.mp4")

function GetMediaCreatedDate($File) {
	$Shell = New-Object -ComObject Shell.Application
	$Folder = $Shell.Namespace($File.DirectoryName)
	$CreatedDate = $Folder.GetDetailsOf($Folder.Parsename($File.Name), 191).Replace([char]8206, ' ').Replace([char]8207, ' ')

	if ($null -ne ($CreatedDate -as [DateTime])) {
		return [DateTime]::Parse($CreatedDate)
	} else {
		return $null
	}
}

function ConvertAsciiArrayToString($CharArray) {
	$ReturnVal = ""
	foreach ($Char in $CharArray) {
		$ReturnVal += [char]$Char
	}
	return $ReturnVal
}

function GetModifiedDate($File) {
	$FileDetail = Get-ChildItem $File
    Write-Host "debut Getmodifiy func" $File
	$DateTimePropertyItem = $FileDetail.LastWriteTime
    #Write-Host "resultat" $DateTimePropertyItem
    	
	$Year = $DateTimePropertyItem.year
    #Write-Host "resultat year "$Year
	$Month = $DateTimePropertyItem.month
	$Day = $DateTimePropertyItem.day
	$Hour = $DateTimePropertyItem.hour
	$Minute = $DateTimePropertyItem.minute
	$Second = $DateTimePropertyItem.second
	
	$DateString = [String]::Format("{0}/{1}/{2} {3}:{4}:{5}", $Year, $Month, $Day, $Hour, $Minute, $Second)
	
	if ($null -ne ($DateString -as [DateTime])) {
		return [DateTime]::Parse($DateString)
	} else {
		return $null
	}
}



function GetCreationDate($File) {
	switch ($File.Extension) { 
        ".mp4" { $CreationDate = GetModifiedDate($File) }
        ".mov" { $CreationDate = GetModifiedDate($File) }  
        ".avi" { $CreationDate = GetModifiedDate($File) } 
        default { $CreationDate = GetModifiedDate($File) }
    }
	return $CreationDate
}

function BuildDesinationPath($Path, $Date) {
	return [String]::Format("{0}\{1}\{2}_{3}", $Path, $Date.Year, $Date.ToString("MM"), $Date.ToString("MMMM"))
}

$RandomGenerator = New-Object System.Random
function BuildNewFilePath($Path, $Date, $Extension) {
	return [String]::Format("{0}\{1}_{2}{3}", $Path, $Date.ToString("yyyy-MM-dd_HH-mm-ss"), $RandomGenerator.Next(100, 1000).ToString(), $Extension)
}

function CreateDirectory($Path){
	if (!(Test-Path $Path)) {
		New-Item $Path -Type Directory
	}
}

function ConfirmContinueProcessing() {
	$Response = Read-Host "Continue? (Y/N)"
	if ($Response.Substring(0,1).ToUpper() -ne "Y") { 
		break 
	}
}

Write-Host "Begin"
$Files = Get-ChildItem $SourceRootPath -Recurse -Include $FileTypesToOrganize
foreach ($File in $Files) {
	$CreationDate = GetCreationDate($File)
	if ($null -ne ($CreationDate -as [DateTime])) {
		$DestinationPath = BuildDesinationPath $DestinationRootPath $CreationDate
		CreateDirectory $DestinationPath
		$NewFilePath = BuildNewFilePath $DestinationPath $CreationDate $File.Extension
		
		Write-Host $File.FullName -> $NewFilePath
		if (!(Test-Path $NewFilePath)) {
			Move-Item $File.FullName $NewFilePath
		} else {
			Write-Host "Unable to rename file. File already exists. "
			#ConfirmContinueProcessing
		}
	} else {
		Write-Host "Unable to determine creation date of file. " $File.FullName
		# ConfirmContinueProcessing
	}
} 

Write-Host "Done"
