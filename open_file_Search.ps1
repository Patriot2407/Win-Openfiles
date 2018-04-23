$filepath = "C:\temp\openfiles.csv"
$dirpath = "C:\temp"
$allextensions = ".pdf .wmf .png .svg .des"
$notdisplayable = ".wmf .png .svg .des"
$fileinfo = "This program shows all Locked (by Write access) and all $allextensions files only. Please note that $notdisplayable files may not display in this list."
$pathinfo = "Also note that the path is relative to the server. B:\Projects is M:\"
if(!(Test-Path -Path $dirpath )){
    New-Item -ItemType directory -Path $dirpath
}
function Get-OpenFile {
    cls
    $primcolor = "Gray"
    $secolor = "Green"
    $search = $args[1]
    $searchpar = $search
    if ($search -eq $null){$searchpar = "All Files"}
    Write-Host "You searched for '$searchpar' on server '$servername'  Here are the results:"
    Write-Host $fileinfo
    Write-Host $pathinfo
    openfiles /query /s $args[0] /fo csv /V | Out-File -Force $filepath
    try{
    $rawtext = Import-CSV $filepath | Select "Accessed By", "Open Mode", "Open File (Path\executable)" | Where-Object {($_."Open File (Path\executable)" -match $search -or $_."Accessed By" -match $search) -and (($_."Open Mode" -match "Write") -or (($_."Open File (Path\executable)" -match ".pdf") -and ($_."Open Mode" -match "Read") -and (!($_."Open File (Path\executable)" -match "D:\\Users")) -and (!($_."Open File (Path\executable)" -match "\\srvsvc")) -and (!($_."Open File (Path\executable)" -match "\\MsFteWds")) -and ($_."Open File (Path\executable)" -match "B:\\Projects" -or $_."Open File (Path\executable)" -match "D:\\Reference" -or $_."Open File (Path\executable)" -match "D:\\Quality System")))} | Format-Table -Wrap -AutoSize | Out-String
    }
    catch{write-host "ERROR DETECTED!" -ForegroundColor Yellow
    write-host "-->$($_.Exception.Message)<--" -ForegroundColor Red
    }
    Remove-Item $filepath
    $prettytext = $rawtext -split "\n"
for ($i=0;$i -lt $prettytext.Length; $i++) {
  if ($i % 2 -eq 0) {
    Write-Host $prettytext[$i] -ForegroundColor $primcolor
  } else {
    Write-Host $prettytext[$i] -ForegroundColor $secolor
  }
}
    Write-Host "Done searching."
   cmd /c pause | out-null
}

function Show-Menu {
param ([string]$Title = 'Main Menu')
     cls
     Write-Host "================ $Title ================" -ForegroundColor green
     Write-Host "1: Press '1' for search by file" -ForegroundColor Magenta
     Write-Host "2: Press '2' for search by user" -ForegroundColor Cyan
     Write-Host "3: Press '3' for all files" -ForegroundColor DarkYellow
     Write-Host "e: Press 'e' to quit this program (or just click the red 'X')" -ForegroundColor DarkRed
}
$servername = "fileserver"
do
                {
                Show-Menu
                $input = Read-Host "Please make a selection"
                 if ($input -eq 1){
                 cls
                 Write-Host "Enter file name or a top level folder. DO NOT USE ANYTHING OTHER THAN ALPHANUMERICAL CHARACTERS! Spaces are allowed."
                 $filename = Read-host -Prompt "File name"
                 $filename = $filename -replace [RegEx]::Escape('\'), "\\"
                 $filename = "$filename"
                 Get-OpenFile $servername $filename}
                 elseif($input -eq 2){
                 cls
                 Write-Host "Enter user name. DO NOT USE ANYTHING OTHER THAN ALPHANUMERICAL CHARACTERS!"
                 $username = Read-host -Prompt "User name"
                 Get-OpenFile $servername $username}
                 elseif($input -eq 3){Get-OpenFile $servername}}
                until ($input -eq 'e')
exit
