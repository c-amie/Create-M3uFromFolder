# -----------------------------------------------------------------------------
# Script: Get-FileMetaDataReturnObject.ps1
# Author: ed wilson, msft
# Date: 01/24/2014 12:30:18
# Keywords: Metadata, Storage, Files
# comments: Uses the Shell.APplication object to get file metadata
# Gets all the metadata and returns a custom PSObject
# it is a bit slow right now, because I need to check all 266 fields
# for each file, and then create a custom object and emit it.
# If used, use a variable to store the returned objects before attempting
# to do any sorting, filtering, and formatting of the output.
# To do a recursive lookup of all metadata on all files, use this type
# of syntax to call the function:
# Get-FileMetaData -folder (gci e:\music -Recurse -Directory).FullName
# note: this MUST point to a folder, and not to a file.
# -----------------------------------------------------------------------------
Function Get-FileMetaData
{
  <#
   .Synopsis
    This function gets file metadata and returns it as a custom PS Object 
   .Description
    This function gets file metadata using the Shell.Application object and
    returns a custom PSObject object that can be sorted, filtered or otherwise
    manipulated.
   .Example
    Get-FileMetaData -folder "e:\music"
    Gets file metadata for all files in the e:\music directory
   .Example
    Get-FileMetaData -folder (gci e:\music -Recurse -Directory).FullName
    This example uses the Get-ChildItem cmdlet to do a recursive lookup of 
    all directories in the e:\music folder and then it goes through and gets
    all of the file metada for all the files in the directories and in the 
    subdirectories.  
   .Example
    Get-FileMetaData -folder "c:\fso","E:\music\Big Boi"
    Gets file metadata from files in both the c:\fso directory and the
    e:\music\big boi directory.
   .Example
    $meta = Get-FileMetaData -folder "E:\music"
    This example gets file metadata from all files in the root of the
    e:\music directory and stores the returned custom objects in a $meta 
    variable for later processing and manipulation.
   .Parameter Folder
    The folder that is parsed for files 
   .Notes
    NAME:  Get-FileMetaData
    AUTHOR: ed wilson, msft
    LASTEDIT: 01/24/2014 14:08:24
    KEYWORDS: Storage, Files, Metadata
    HSG: HSG-2-5-14
   .Link
     Http://www.ScriptingGuys.com
 #Requires -Version 2.0
 #>
 Param( [string[]]$folder )
 foreach( $sFolder in $folder )
  {
   $a = 0
   $objShell = New-Object -ComObject Shell.Application
   $objFolder = $objShell.namespace( $sFolder )

   foreach ( $File in $objFolder.items() )
    { 
     $FileMetaData = New-Object PSOBJECT
      for ( $a ; $a  -le 266; $a++ )
       { 
         if( $objFolder.getDetailsOf( $File, $a ))
           {
             $hash += @{ $( $objFolder.getDetailsOf( $objFolder.items, $a ))  =
                   $( $objFolder.getDetailsOf( $File, $a )) }
            $FileMetaData | Add-Member $hash
            $hash.clear() 
           } #end if
       } #end for 
     $a=0
     $FileMetaData
    } #end foreach $file
  } #end foreach $sfolder
} #end Get-FileMetaData


<#
    .SYNOPSIS
    Generates a M3U Playlist file from a folder of media files

    .DESCRIPTION
    Uses ID3 tag information to create and populate a basic M3U playlist file based upon the media content of a given source folder.
    The output filename of the M3U file will be set to match the value of <Title>.m3u
    Note: Requires that the MP3 files have properly formatted embedded ID3 tags

    .PARAMETER SourceFolder
    Required. The folder source containing the MP3 files to scan (flat files only, no folder recursion is used)

    .PARAMETER Title
    The name of the book formatted using File System Safe Characters (i.e. not: * " / \ < > : | ?)

    .PARAMETER Year
    The year that the book was published in

    .PARAMETER Artist
    The name of the album artist

    .PARAMETER Genre
    The name of the genre of the album

    .PARAMETER Thuumbnail
    The filename of a PNG or JPG file located in the SourceFolder path to use as the M3U album art

    .PARAMETER ThuumbnailUrl
    Fully qualifid http or https path to the Thumbnail image

    .EXAMPLE
    PS> Create-M3uFromFolder -SourceFolder 'C:\MyFolder' -Title 'The Wind in the Willows' -Year 1908 -Thuumbnail 'thumbnail.jpg'
    
    .NOTES
    NAME:     Create-M3uFromFolder
    AUTHOR:   C:Amie
    LASTEDIT: 2025/04/19 16:07:23
    KEYWORDS: Storage, Files, Metadata, MP3, M3U
    
    .LINK
    https://www.c-amie.co.uk/
 #Requires -Version 2.0
#>
function Create-M3uFromFolder {
  param (
    [Parameter(Mandatory=$true)]
      [string]$SourceFolder,
    [Parameter(Mandatory=$true)]
      [string]$Title,
    [Parameter(Mandatory=$true)]
      [int]$Year,
    [Parameter(Mandatory=$false)]
      [string]$Artist,
    [Parameter(Mandatory=$false)]
      [string]$Genre,
    [Parameter(Mandatory=$false)]
      [string]$Thuumbnail,
    [Parameter(Mandatory=$false)]
      [string]$ThuumbnailUrl
  )
  if (-Not $(Test-Path $SourceFolder -PathType Container)) {
    Write-Host "The path '$SourcePath' does not exist" -ForegroundColor Red
    Exit 1
  }

  Write-Host " Processing Book: $Title ($Year)" -ForegroundColor Green

  $strOut = "#EXTM3U`r`n"
  $strOut += "#EXTALB:$($Title) ($($Year))`r`n"
  if ($Artist) {
    $strOut += "#EXTART:$($Artist)`r`n"
  }
  if ($Genre) {
    $strOut += "#EXTGENRE:$($Genre)`r`n"
  }
  if ($ThuumbnailUrl) {
    $strOut += "#EXTALBUMARTURL:$($ThuumbnailUrl)`r`n"
  }
  Get-FileMetaData -folder $SourceFolder | % {
  $mp3 = $_
    if ($mp3.Kind -eq 'Music') {
      Write-Host " - Adding: $($mp3.Title)" -ForegroundColor Green
      $strOut += "#EXTINF:$([TimeSpan]::Parse($mp3.Length).TotalSeconds) "
      if ($Thuumbnail) {
        $strOut += "logo=`"$($Thuumbnail)`""
      }
      $strOut += ",$($mp3.Title)`r`n$($mp3.Name)`r`n"
    }
  }
  
  #$strOut | Out-File -FilePath "$($strSource)\$($Title).m3u" -Encoding utf8   # Writes a BOM which won't work in UTF-8 M3U files
  [System.IO.File]::WriteAllLines("$($SourceFolder)\$($Title).m3u", $strOut)   # Does not write a BOM

  Write-Host " Done!" -ForegroundColor Green
}
