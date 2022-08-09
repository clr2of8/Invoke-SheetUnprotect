function Invoke-SheetUnprotect{
<#
  .SYNOPSIS
    This module is used to set the password for protected sheets in a Microsoft Excel Document to xyz
    Author: Carrie Roberts (@OrOneEqualsOne)
    License: BSD 3-Clause
    Dependencies:
    Version: 1.0

  .DESCRIPTION
    This module is used to set the password for protected sheets in a Microsoft Excel Document to xyz

  .PARAMETER OfficeFile
    Name of source MS OFfice document to modify sheet protection password on.


  .EXAMPLE

    C:\PS> Invoke-SheetUnprotect  .\Protected.xlsx

    Description
    -----------
    This command will create a copy of the Proctect.xlsx file with the password for any protected sheet set to xyz. The file will be 
    saved to the same directory and have "-xyz" appended to the file name. e.g. Protected-xyz.xlsx
    
#>

  Param
  (
    [Parameter(Position = 0, Mandatory = $true)]
    [string]
    $OfficeFile = "" 

)

Write-Host -ForegroundColor Yellow -NoNewline "Workin' it "

# Copy office document to temp dir
$fnwoe = [System.IO.Path]::GetFileNameWithoutExtension($OfficeFile)
$zipFile = (Join-Path $env:Temp $fnwoe) + ".zip"
Copy-Item -Path $OfficeFile -Destination $zipFile -Force

#unzip MS Office document to temporary location
$Destination = Join-Path $env:TEMP $fnwoe
Expand-ZIPFile $zipFile $Destination

# remove sheetprotection from each sheet
$DocPropFolder = Join-Path $Destination "xl" | Join-Path -ChildPath "worksheets"
gci $DocPropFolder -filter *.xml |
ForEach-Object{
	Remove-SheetProtection $_.FullName
}

#zip files back up with MS Office extension
$zipfileName = $Destination + ".zip"
Create-ZIPFile $Destination $zipfileName

#copy zip file back to original $OfficeFile location and rename with an appended "-xyz" and the original extension
$newOfficeFileName = Join-Path ([System.IO.Path]::GetDirectoryName($OfficeFile)) ([System.IO.Path]::GetFileNameWithoutExtension($OfficeFile) + "-xyz" + [System.IO.Path]::GetExtension($OfficeFile))
Copy-Item $zipfileName $newOfficeFileName
Write-Host -ForegroundColor Green "`rThe new file with added comment has been written to $newOfficeFileName.`nDONE!"
}

function Expand-ZIPFile($file, $destination)
{
    #delete the destination folder if it already exists
    If(test-path $destination)
    {
        Remove-Item -Recurse -Force $destination
    }
    New-Item -ItemType Directory -Force -Path $destination | Out-Null

    
    #extract to the destination folder
    $shell = new-object -com shell.application
    $zip = $shell.NameSpace($file)
    $shell.namespace($destination).copyhere($zip.items())
}

#Zip code is from https://serverfault.com/questions/456095/zipping-only-files-using-powershell
function Create-ZIPFile($folder, $zipfileName)
{
    #delete the zip file if it already exists
    If(test-path $zipfileName)
    {
        Remove-Item -Force $zipfileName
    }
    set-content $zipfileName ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
    (dir $zipfileName).IsReadOnly = $false  

    $shellApplication = new-object -com shell.application
    $zipPackage = $shellApplication.NameSpace($zipfileName)

    $files = Get-ChildItem -Path $folder
    foreach($file in $files) 
    { 
            $zipPackage.CopyHere($file.FullName)
            #using this method, sometimes files can be 'skipped'
            #this 'while' loop checks each file is added before moving to the next
            while($zipPackage.Items().Item($file.name) -eq $null){
                Write-Host -ForegroundColor Yellow -NoNewline ". "
                Start-sleep -seconds 1
            }
    }
}

function Remove-SheetProtection($DocPropFile)
{

   $xmlDoc = [System.Xml.XmlDocument](Get-Content $DocPropFile);

    Try{
        #overwrite the password values with the hashes representing a password of xyz
        $xmlDoc.worksheet.sheetProtection.hashValue = "cyHPzii8CCjh7yMMdBXlsICwP7PpPB7bP2UxJYvhD2hHF2onWlRGcKZx37WFTdIzZTKZcH5NpJ5voZ3tGmgkMA=="
        $xmlDoc.worksheet.sheetProtection.saltValue = "5gCuGkmzk7/E3/HJnumPWA=="
    }
    Catch {
    }

   $xmlDoc.Save($DocPropFile)
}