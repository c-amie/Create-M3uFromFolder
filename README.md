# Create-M3uFromFolder
PowerShell function to create a M3U playlist file from a folder of MP3s

Uses Get-FileMetaDataReturnObject by Ed Wilson (https://github.com/mattlite/powershell/blob/master/Get-FileMetaDataReturnObject.ps1) to gather MP3 meta data which is the used to create a basic M3U playlist file basead upon ID3 tag information and user supplied values.

# Usage Example
```Create-M3uFromFolder -SourceFolder 'C:\MyFolder' -Title 'The Wind in the Willows' -Year 1908 -Artist 'Kenneth Grahame' -Genre 'Audio Book' -Thuumbnail 'thumbnail.jpg'```
