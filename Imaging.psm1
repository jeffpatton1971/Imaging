[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | out-null

function ConvertTo-Bitmap {
    <#
    .Synopsis
    Converts an image to a bitmap (.bmp) file.

    .Description
    The ConvertTo-Bitmap function converts image files to .bmp file format.
    You can specify the desired image quality on a scale of 1 to 100.

    ConvertTo-Bitmap takes only COM-based image objects of the type that Get-Image returns.
    To use this function, use the Get-Image function to create image objects for the image files, 
    then submit the image objects to ConvertTo-Bitmap.

    The converted files have the same name and location as the original files but with a .bmp file name extension. 
    If a file with the same name already exists in the location, ConvertTo-Bitmap declares an error. 

    .Parameter Image
    Specifies the image objects to convert.
    The objects must be of the type that the Get-Image function returns.
    Enter a variable that contains the image objects or a command that gets the image objects, such as a Get-Image command.
    This parameter is optional, but if you do not include it, ConvertTo-Bitmap has no effect.

    .Parameter Quality
    A number from 1 to 100 that indicates the desired quality of the .bmp file.
    The default is 100, which represents the best possible quality.

    .Parameter HideProgress
    Hides the progress bar.

    .Parameter Remove
    Deletes the original file. By default, both the original file and new .bmp file are saved. 

    .Notes
    ConvertTo-Bitmap uses the Windows Image Acquisition (WIA) Layer to convert files.

    .Link
    "Image Manipulation in PowerShell": http://blogs.msdn.com/powershell/archive/2009/03/31/image-manipulation-in-powershell.aspx

    .Link
    "ImageProcess object": http://msdn.microsoft.com/en-us/library/ms630507(VS.85).aspx

    .Link 
    Get-Image

    .Link
    ConvertTo-JPEG

    .Example
    Get-Image .\MyPhoto.png | ConvertTo-Bitmap

    .Example
    # Deletes the original BMP files.
    dir .\*.jpg | get-image | ConvertTo-Bitmap –quality 100 –remove -hideProgress

    .Example
    $photos = dir $home\Pictures\Vacation\* -recurse –include *.jpg, *.png, *.gif
    $photos | get-image | ConvertTo-Bitmap
    #>
    param(
    [Parameter(ValueFromPipeline=$true)]    
    $Image,
    
    [ValidateRange(1,100)]
    [int]$Quality = 100
    )
    process {
     if (($image -is [String]) -or ($image -is [System.io.FileInfo])) {Get-Image $image | convertTo-Bitmap -quality $quality ;return}
        if  ($image.count -gt 1) {$image | convertTo-Bitmap -quality $quality ; return}
        if  (-not $image.Loadfile -and -not $image.Fullname) { return }
        write-verbose ("Processing $($image.fullName)")
        $noExtension = $image.Fullname -replace "\.\w*$",""   # "\.\w*$" means dot followed by any number of alpha chars, followed by end of string - i.e file extension
        $process = New-Object -ComObject Wia.ImageProcess
        $convertFilter = $process.FilterInfos.Item("Convert").FilterId
        $process.Filters.Add($convertFilter)
        $process.Filters.Item(1).Properties.Item("Quality") = $quality
        $process.Filters.Item(1).Properties.Item("FormatID") = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
        $newImg = $process.Apply($image.PSObject.BaseObject)
        $newImg.SaveFile("$noExtension.bmp")
    }
}
function ConvertTo-Jpeg {
<#
        .Synopsis
            Converts a file to a JPG of the specified quality in the same folder
        .Description
            Converts a file to a JPG of the specified quality in the same folder. 
            If the file is already a JPG it will be overwritten at the new quality setting
        .Example
            C:\PS>  Dir -recure -include *.tif | Convert-toJPeg .\myImage.bmp
            Creates creates JPG images of quality 100 for all tif files in the current directory and it's sub directories
        .Example
            C:\PS>  Dir -recure -include *.tif | Convert-toJPeg -quality 75
            Creates JPG images of quality 75 for all tif files in the current directory and it's sub directories
        .Parameter Image
            An image object, a path to an image, or a file object representing an image file. It may be passed via the pipeline.
        .Parameter Quality
            Range 1-100, sets image quality (100 highest), lower quality will use higher rates of compression.
            The default is 100. 
    #>
[CmdletBinding()]
    param(
    [Parameter(ValueFromPipeline=$true)]    
    $image,
    
    [ValidateRange(1,100)]
    [int]$quality = 100
    )
    process {
        if (($image -is [String]) -or ($image -is [System.io.FileInfo])) {Get-Image $image | convertTo-Jpeg -quality $quality ;return}
        if  ($image.count -gt 1) {$image | convertTo-Jpeg -quality $quality ; return}
        if  (-not $image.Loadfile -and -not $image.Fullname) { return }
        write-verbose ("Processing $($image.fullName)")
        $noExtension = $image.Fullname -replace "\.\w*$",""   # "\.\w*$" means dot followed by any number of alpha chars, followed by end of string - i.e file extension
        $process = New-Object -ComObject Wia.ImageProcess
        $convertFilter = $process.FilterInfos.Item("Convert").FilterId
        $process.Filters.Add($convertFilter)
        $process.Filters.Item(1).Properties.Item("Quality") = $quality
        $process.Filters.Item(1).Properties.Item("FormatID") = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
        $newImg = $process.Apply($image.PSObject.BaseObject)
        $newImg.SaveFile("$noExtension.jpg")    
    }
}
Function Copy-Image {
<#
        .Synopsis
            Copies an image, applying EXIF data from GPS data points   
        .Description
            Copies an image, applying EXIF data from GPS data points   
        .Example
            C:\PS>  Dir E:\dcim –inc IMG*.jpg –rec | Copy-Image -Keywords "Oxfordshire" -rotate -DestPath "$env:userprofile\pictures\oxford" -replace  "IMG","OX-"
            Copies IMG files from folders under E:\DCIM to the user's picture\Oxford folder, replacing IMG in the file name with OX-.
            The Keywords field is set to Oxfordshire, pictures are GeoTagged with the data in $points and rotated. 
        .Parameter Image
            A WIA image object, a path to an image, or a file object representing an image file. It may be passed via the pipeline.
        .Parameter Destination
            The FOLDER to which the file should be saved.
        .Parameter Keywords
            If specified, sets the keywords Exif field.
        .Parameter Title
            If specified, sets the Title Exif field..    
        .Parameter Replace
            If specified, this contains two values seperated by a comma specifying a replacement in the file name
        .Parameter Rotate
            If this switch is specified, the image will be auto-rotated based on its orientation filed
        .Parameter NoClobber
            Unless this switch is specified, a pre-existing image WILL be over-written
        .Parameter ReturnInfo
            If this switch is specified, the path to the saved image will be returned. 
    #>
[CmdletBinding(SupportsShouldProcess=$true)]
Param ( [Parameter(ValueFromPipeline=$true, Mandatory=$true)][Alias("Path","FullName")]$image , 
        [Parameter(Mandatory=$true)][ValidateScript({Test-path $_ })][string]$Destination ,  
        $keywords , $Title, $replace,$filter,[switch]$Rotate,[switch]$NoClobber,[switch]$ReturnInfo, $psc 
)
process {
        if ($psc -eq $null)  {$psc = $pscmdlet} ; if (-not $PSBoundParameters.psc) {$PSBoundParameters.add("psc",$psc)}
        if ($image -is [system.io.fileinfo] ) {$image = $image.FullName }
        if ($image -is [String]             ) {[Void]$PSBoundParameters.Remove("Image") 
                                               Get-Image $image | Copy-Image @PSBoundParameters
                                               return
        }
        if ($Image.count -gt 1              ) {[Void]$PSBoundParameters.Remove("Image") 
                                               $Image | ForEach-object {Copy-Image -image $_ @PSBoundParameters}
                                               return
        }
        if ($image -is [__comObject])  {
           Write-Verbose ("Processing " + $image.fullname)
           if (-not $filter)  {$filter = new-Imagefilter}
           if ($rotate)       {$orient=Get-ExifItem -image  $image       -ExifID $ExifIDOrientation}  # Leave $orient unset if we aren't rotating
           if ($keywords)     {Add-exifFilter       -filter $filter      -ExifID $ExifIDKeywords   -typeid 1101 -string $keywords }
           if ($Title)        {Add-exifFilter       -filter $filter      -ExifID $ExifIDTitle      -typeid 1101 -string $Title    }
           if ($orient -eq 8) {Add-RotateFlipFilter -filter $filter      -angle  270   # Orientation 8=90 degrees, 6=270 degrees, rotate round to 360
                               Add-exifFilter       -filter $filter      -ExifID $ExifIDOrientation -typeid $1003 -value 1      
                               write-verbose "Rotating image counter-clockwise"}
           if ($orient -eq 6) {Add-RotateFlipFilter -filter $filter      -angle  90  
                               Add-exifFilter       -filter $filter      -ExifID $ExifIDOrientation -typeid $1003 -value 1      
                               write-verbose "Rotating image clockwise"}
           if ($replace)      {$SavePath= join-path -Path   (Resolve-Path $Destination) -ChildPath ((Split-Path $image.FullName -Leaf) -Replace $replace)}
           else               {$SavePath= join-path -Path   (Resolve-Path $Destination) -ChildPath  (Split-Path $image.FullName -Leaf)  }
           Set-ImageFilter    -image $image         -filter $filter      -SaveName $savePath -noClobber:$NoClobber -psc $psc
           $orient = $image =  $filter = $null
           if ($returnInfo) {$SavePath}
        }
    }
}
function Get-Image {
    <#
        .Synopsis
            Returns an image object for a file
        .Description
            Uses the Windows Image Acquisition COM object to get image data
        .Example
            Get-ChildItem $env:UserProfile\Pictures -Recurse | Get-Image        
        .Parameter file
            The file to get an image from
    #>
    param(    
    [Parameter(ValueFromPipelineByPropertyName=$true,Mandatory=$true)]
    [Alias('FullName',"FileName")]
    [ValidateScript({Test-path $_ })][string]$Path)
    
    process {
        foreach ($file in (resolve-path -Path $path) ) {
            $image  = New-Object -ComObject Wia.ImageFile        
            try {        
                Write-Verbose "Loading file $($realItem.FullName)"
                $image.LoadFile($file.path)
                $image | 
                    Add-Member NoteProperty FullName $File -PassThru | 
                    Add-Member ScriptMethod Resize {
                        param($width, $height, [switch]$DoNotPreserveAspectRatio)                    
                        $image = New-Object -ComObject Wia.ImageFile
                        $image.LoadFile($this.FullName)
                        $filter = Add-ScaleFilter @psBoundParameters -passThru -image $image
                        $image = $image | Set-ImageFilter -filter $filter -passThru
                        Remove-Item $this.Fullname
                        $image.SaveFile($this.FullName)                    
                    } -PassThru | 
                    Add-Member ScriptMethod Crop {
                        param([Double]$left, [Double]$top, [Double]$right, [Double]$bottom)
                        $image = New-Object -ComObject Wia.ImageFile
                        $image.LoadFile($this.FullName)
                        $filter = Add-CropFilter @psBoundParameters -passThru -image $image
                        $image = $image | Set-ImageFilter -filter $filter -passThru
                        Remove-Item $this.Fullname
                        $image.SaveFile($this.FullName)                    
                    } -PassThru | 
                    Add-Member ScriptMethod FlipVertical {
                        $image = New-Object -ComObject Wia.ImageFile
                        $image.LoadFile($this.FullName)
                        $filter = Add-RotateFlipFilter -flipVertical -passThru 
                        $image = $image | Set-ImageFilter -filter $filter -passThru
                        Remove-Item $this.Fullname
                        $image.SaveFile($this.FullName)                    
                    } -PassThru | 
                    Add-Member ScriptMethod FlipHorizontal {
                        $image = New-Object -ComObject Wia.ImageFile
                        $image.LoadFile($this.FullName)
                        $filter = Add-RotateFlipFilter -flipHorizontal -passThru 
                        $image = $image | Set-ImageFilter -filter $filter -passThru
                        Remove-Item $this.Fullname
                        $image.SaveFile($this.FullName)                    
                    } -PassThru |
                    Add-Member ScriptMethod RotateClockwise {
                        $image = New-Object -ComObject Wia.ImageFile
                        $image.LoadFile($this.FullName)
                        $filter = Add-RotateFlipFilter -angle 90 -passThru 
                        $image = $image | Set-ImageFilter -filter $filter -passThru
                        Remove-Item $this.Fullname
                        $image.SaveFile($this.FullName)                    
                    } -PassThru |
                    Add-Member ScriptMethod RotateCounterClockwise {
                        $image = New-Object -ComObject Wia.ImageFile
                        $image.LoadFile($this.FullName)
                        $filter = Add-RotateFlipFilter -angle 270 -passThru 
                        $image = $image | Set-ImageFilter -filter $filter -passThru
                        Remove-Item $this.Fullname
                        $image.SaveFile($this.FullName)                    
                    } -PassThru 
                    
            } catch {
                Write-Verbose $_
            }
        }     
    }    
}
Function Get-IndexedItem {
    <#
       .SYNOPSIS
            Gets files which have been indexed by Windows desktop search
       .Description
            Searches the Windows index on the local computer or a remote file serving computer
            Looking for file properties or free text searching over contents       
        .PARAMETER Filter
            Alias WHERE, INCLUDE
            A single string containing a WHERE condition, or multiple conditions linked with AND
            or Multiple strings each with a single Condition, which will be joined together.
            The function tries to add Prefixes and single quotes if they are omitted
            If no =, >,< , Like or Contains is specified the terms will be used in a freeText contains search
            Syntax Information for CONTAINS and REETEXT can be found at 
            http://msdn.microsoft.com/en-us/library/dd626247(v=office.11).aspx
        .PARAMETER OrderBy
            Alias SORT
            Either a single string containing one or more Order BY conditions, 
            or multiple string each with a single condition which will be joined together            
        .PARAMETER Path
            A single string containing a path which should be searched. 
            This may be a UNC path to a share on a remote computer 
        .PARAMETER First
            Alias TOP
            A single integer representing the number of items to be returned. 
        .PARAMETER Value
            Alias GROUP
            A single string containing a Field name. 
            If specified the search will return the Values in this field, instead of objects
            for the items found by the query terms. 
        .PARAMETER Recurse
            If Path is specified only a single folder is searched Unless -Recurse is specified
            If path is not specified the whole index is searched, and recurse is ignored. 
        .PARAMETER List
            Instead of querying the index produces a list of known field names, with short names and aliases
            which may be used instead.
        .PARAMETER NoFiles
            Normally if files are found the command returns a file object with additional properties,
            which can be piped into commands which accept files. This switch prevents the file being fetched
            improving performance when the file object is not needed. 
        .EXAMPLE
            Get-IndexedItem -Filter "Contains(*,'Stingray')", "kind = 'picture'", "keywords='portfolio'" 
            Finds picture files anywhere on the local machine, which have 'Portfolio' as a keyword tag,
            and 'stringray' in any indexed property.
        .EXAMPLE
            Get-IndexedItem Stingray, kind=picture, keyword=portfolio | copy -destination e:\
            Finds the same pictures as the previous example but uses Keyword as a alias for KeywordS, and
            leaves the ' marks round Portfolio and Contains() round stingray to be automatically inserted  
            Copies the found files to drive E: 
        .EXAMPLE
            Get-IndexedItem -filter stingray -path OneIndex14:// -recurse    
            Finds OneNote items containing "Stingray" (note, nothing will be found without -recurse) 
        .EXAMPLE
            start (Get-IndexedItem -filter stingray -path OneIndex14:// -recurse -first 1 -orderby rank)
            Finds the highest ranked one not page for stingray and opens it. 
            Note Start-process (canonical name for Start) does not support piped input. 
        .EXAMPLE
            Get-IndexedItem -filter stingray -path ([system.environment]::GetFolderPath( [system.environment+specialFolder]::MyPictures )) -recurse    
            Looks for pictures with stingray in any indexed property, limiting the scope of the search 
            to the current users 'My Pictures' folder and its subfolders.
        .EXAMPLE
            Get-IndexedItem -Filter "system.kind = 'recordedTV' " -order "System.RecordedTV.RecordingTime" -path "\\atom-engine\users" -recurse | format-list path,title,episodeName,programDescription
            Finds recorded TV files on a remote server named 'Atom-Engine' which are accessible via a share named 'users'. 
            Field name prefixes are specified explicitly instead of letting the function add them
            Results are displayed as a list using a subset of the available fields specific to recorded TV
        .EXAMPLE
            Get-IndexedItem -Value "kind" -path \\atom-engine\users  -recurse
            Lists the kinds of files available on the on the 'users' share of a remote server named 'Atom-Engine'
        .EXAMPLE    
            Get-IndexedItem -Value "title" -filter "kind=recordedtv" -path \\atom-engine\users  -recurse
            Lists the titles of RecordedTv files available on the on the 'users' share of a remote server named 'Atom-Engine'
        .EXAMPLE
           Start (Get-IndexedItem -path "\\atom-engine\users" -recurse -Filter "title= 'Formula 1' " -order "System.RecordedTV.RecordingTime DESC" -top 1 )    
           Finds files entitled "Formula 1" on the 'users' share of a remote server named 'Atom-Engine'
           Selects the most recent one by TV recording date, and opens it on the local computer. 
           Note: start does not support piped input. 
        .EXAMPLE
           Get-IndexedItem -Filter "System.Kind = 'Music' AND AlbumArtist like '%'  " | Group-Object -NoElement -Property "AlbumArtist" | sort -Descending -property count
           Gets all music files with an Album Artist set, using a single combined where condition and a mixture 
           of implicit and explicit field prefixes.  
           The result is grouped by Artist and sorted to give popular artist first
        .EXAMPLE
           Get-IndexedItem -path c:\ -recurse  -Filter cameramaker=pentax* -Property focallength | group focallength -no | sort -property @{e={[double]$_.name}}   
           Gets all the items which have a the camera maker set to pentax, anywhere on the C: driv
           but ONLY get thier focallength property, and return a sorted count of how many of each focal length there are. 
    #>
    #$t=(Get-IndexedItem -Value "title" -filter "kind=recordedtv" -path \\atom-engine\users  -recurse | Select-List -Property title).title
    #start (Get-IndexedItem -filter "kind=recordedtv","title='$t'" -path \\atom-engine\users  -recurse | Select-List -Property ORIGINALBROADCASTDATE,PROGRAMDESCRIPTION)
[CmdletBinding()]
Param ( [Alias("Where","Include")][String[]]$Filter , 
        [String]$path, 
        [Alias("Sort")][String[]]$orderby, 
        [Alias("Top")][int]$First,
        [Alias("Group")][String]$Value, 
        [Alias("Select")][String[]]$Property, 
        [Switch]$recurse,
        [Switch]$list,
        [Switch]$NoFiles)

#Alias definitions take the form  AliasName = "Full.Cannonical.Name" ; 
#Any defined here will be accepted as input field names in -filter and -OrderBy parameters
#and will be added to output objects as AliasProperties. 
 $PropertyAliases   = @{Width         ="System.Image.HorizontalSize"; Height        = "System.Image.VerticalSize";  Name    = "System.FileName" ; 
                        Extension     ="System.FileExtension"       ; CreationTime  = "System.DateCreated"       ;  Length  = "System.Size" ; 
                        LastWriteTime ="System.DateModified"        ; Keyword       = "System.Keywords"          ;  Tag     = "System.Keywords"
                        CameraMaker  = "System.Photo.Cameramanufacturer"}

 $fieldTypes = "System","Photo","Image","Music","Media","RecordedTv","Search" 
#For each of the field types listed above, define a prefix & a list of fields, formatted as "Bare_fieldName1|Bare_fieldName2|Bare_fieldName3"
#Anything which appears in FieldTypes must have a prefix and fields definition. 
#Any definitions which don't appear in fields types will be ignored 
#See http://msdn.microsoft.com/en-us/library/dd561977(v=VS.85).aspx for property info.  
 
 $SystemPrefix     = "System."            ;     $SystemFields = "ItemName|ItemUrl|FileExtension|FileName|FileAttributes|FileOwner|ItemType|ItemTypeText|KindText|Kind|MIMEType|Size|DateModified|DateAccessed|DateImported|DateAcquired|DateCreated|Author|Company|Copyright|Subject|Title|Keywords|Comment|SoftwareUsed"
 $PhotoPrefix      = "System.Photo."      ;      $PhotoFields = "fNumber|ExposureTime|FocalLength|IsoSpeed|PeopleNames|DateTaken|Cameramodel|Cameramanufacturer|orientation"
 $ImagePrefix      = "System.Image."      ;      $ImageFields = "Dimensions|HorizontalSize|VerticalSize"
 $MusicPrefix      = "System.Music."      ;      $MusicFields = "AlbumArtist|AlbumID|AlbumTitle|Artist|BeatsPerMinute|Composer|Conductor|DisplayArtist|Genre|PartOfSet|TrackNumber"
 $MediaPrefix      = "System.Media."      ;      $MediaFields = "Duration"
 $RecordedTVPrefix = "System.RecordedTV." ; $RecordedTVFields = "ChannelNumber|EpisodeName|OriginalBroadcastDate|ProgramDescription|RecordingTime|StationName"
 $SearchPrefix     = "System.Search."     ;     $SearchFields = "AutoSummary|HitCount|Rank|Store"
 
 if ($list)  {  #Output a list of the fields and aliases we currently support. 
    $( foreach ($type in $fieldTypes) { 
          (get-variable "$($type)Fields").value -split "\|" | select-object @{n="FullName" ;e={(get-variable "$($type)prefix").value+$_}},
                                                                            @{n="ShortName";e={$_}}    
       }
    ) + ($PropertyAliases.keys | Select-Object  @{name="FullName" ;expression={$PropertyAliases[$_]}},
                                                @{name="ShortName";expression={$_}}
    ) | Sort-Object -Property @{e={$_.FullName -split "\.\w+$"}},"FullName" 
  return
 }  
  
#Make a giant SELECT clause from the field lists; replace "|" with ", " - field prefixes will be inserted later.
#There is an extra comma to ensure the last field name is recognized and gets a prefix. This is tidied up later
 if ($first)    {$SQL =  "SELECT TOP $first "}
 else           {$SQL =  "SELECT "}
 if ($property) {$SQL += ($property -join ", ") + ", "}
 else {
    foreach ($type in $fieldTypes) { 
        $SQL += ((get-variable "$($type)Fields").value -replace "\|",", " ) + ", " 
    }
 }   
  
#IF a UNC name was specified as the path, build the FROM ... WHERE clause to include the computer name.
 if ($path -match "\\\\([^\\]+)\\.") {
       $sql += " FROM $($matches[1]).SYSTEMINDEX WHERE "  
 } 
 else {$sql += " FROM SYSTEMINDEX WHERE "} 
 
#If a WHERE condidtion was provided via -Filter, add it now   

 if ($Filter) { #Convert * to % 
                $Filter = $Filter -replace "(?<=\w)\*","%"
                #Insert quotes where needed any condition specified as "keywords=stingray" is turned into "Keywords = 'stingray' "
                $Filter = $Filter -replace "\s*(=|<|>|like)\s*([^\''\d][^\d\s\'']*)$"  , ' $1 ''$2'' '
                # Convert "= 'wildcard'" to "LIKE 'wildcard'" 
                $Filter = $Filter -replace "\s*=\s*(?='.+%'\s*$)" ," LIKE " 
                #If a no predicate was specified, use the term in a contains search over all fields.
                $filter = ($filter | ForEach-Object {
                                if ($_ -match "'|=|<|>|like|contains|freetext") {$_}
                                else {"Contains(*,'$_')"}
                }) 
                #if $filter is an array of single conditions join them together with AND 
                  $SQL += $Filter -join " AND "  } 
                  
 #If a path was given add SCOPE or DIRECTORY to WHERE depending on whether -recurse was specified. 
 if ($path)     {if ($path -notmatch "\w{4}:") {$path = "file:" + (resolve-path -path $path).providerPath}  # Path has to be in the form "file:C:/users" 
                $path  = $path -replace "\\","/"
                if ($sql -notmatch "WHERE\s$") {$sql += " AND " }                       #If the SQL statement doesn't end with "WHERE", add "AND"  
                if ($recurse)                  {$sql += " SCOPE = '$path' "       }     #INDEX uses SCOPE <folder> for recursive search, 
                else                           {$sql += " DIRECTORY = '$path' "   }     # and DIRECTORY <folder> for non-recursive
 }   
 
 if ($Value) {
                if ($sql -notmatch "WHERE\s$") {$sql += " AND " }                       #If the SQL statement doesn't end with "WHERE", add "AND"  
                                                $sql += " $Value Like '%'" 
                                                $sql =  $SQL -replace "^SELECT.*?FROM","SELECT $Value, FROM"
 }
 
 #If the SQL statement Still ends with "WHERE" we'd return everything in the index. Bail out instead  
 if ($sql -match "WHERE\s*$")  { Write-warning "You need to specify either a path , or a filter." ; return} 
 
 #Add any order-by condition(s). Note there is an extra trailing comma to ensure field names are recognised when prefixes are inserted . 
 if ($Value) {$SQL =  "GROUP ON $Value, OVER ( $SQL )"}
 elseif ($orderby)  {$sql += " ORDER BY " + ($orderby   -join " , " ) + ","}             
 
 # For each entry in the PROPERTYALIASES Hash table look for the KEY part being used as a field name
 # and replace it with the associated value. The operation becomes
 # $SQL  -replace "(?<=\s)CreationTime(?=\s*(=|\>|\<|,|Like))","System.DateCreated" 
 # This translates to "Look for 'CreationTime' preceeded by a space and followed by ( optionally ) some spaces, and then
 # any of '=', '>' , '<', ',' or 'Like' (Looking for these prevents matching if the word is a search term, rather than a field name)
 # If you find it, replace it with "System.DateCreated" 
 
 $PropertyAliases.Keys | ForEach-Object { $sql= $SQL -replace "(?<=\s)$($_)(?=\s*(=|>|<|,|Like))",$PropertyAliases.$_}      

 # Now a similar process for all the field prefixes: this time the regular expression becomes for example,
 # $SQL -replace "(?<!\s)(?=(Dimensions|HorizontalSize|VerticalSize))","System.Image." 
 # This translates to: "Look for a place which is preceeded by space and  followed by 'Dimensions' or 'HorizontalSize'
 # just select the place (unlike aliases, don't select the fieldname here) and put the prefix at that point.  
 foreach ($type in $fieldTypes) { 
    $fields = (get-variable "$($type)Fields").value 
    $prefix = (get-variable "$($type)Prefix").value 
    $sql = $sql -replace "(?<=\s)(?=($Fields)\s*(=|>|<|,|Like))" , $Prefix
 }
 
 # Some commas were  put in just to ensure all the field names were found but need to be removed or the SQL won't run
 $sql = $sql -replace "\s*,\s*FROM\s+" , " FROM " 
 $sql = $sql -replace "\s*,\s*OVER\s+" , " OVER " 
 $sql = $sql -replace "\s*,\s*$"       , "" 
 
 #Finally we get to run the query: result comes back in a dataSet with 1 or more Datatables. Process each dataRow in the first (only) table
 write-debug $sql 
 $adapter = new-object system.data.oledb.oleDBDataadapter -argumentlist $sql, "Provider=Search.CollatorDSO;Extended Properties=’Application=Windows’;"
 $ds      = new-object system.data.dataset
 if ($adapter.Fill($ds)) { foreach ($row in $ds.Tables[0])  {
    #If the dataRow refers to a file output a file obj with extra properties, otherwise output a PSobject
    if ($Value) {$row | Select-Object -Property @{name=$Value; expression={$_.($ds.Tables[0].columns[0].columnname)}}}
    else {
        if (($row."System.ItemUrl" -match "^file:") -and (-not $NoFiles)) { 
               $obj = (Get-item -force -Path (($row."System.ItemUrl" -replace "^file:","") -replace "\/","\"))
        }
        else { 
               if ($row."System.ItemUrl") {
                     $obj = New-Object psobject -Property @{Path = $row."System.ItemUrl"}
                     Add-Member -force -InputObject $obj -Name "ToString"  -MemberType "scriptmethod" -Value {$this.path} 
               }
               else {$obj = New-Object psobject }   
        }
        if ($obj) {
            #Add all the the non-null dbColumns removing the prefix from the property name. 
            foreach ($prop in (Get-Member -InputObject $row -MemberType property | where-object {$row."$($_.name)" -isnot [system.dbnull] })) {                            
                Add-member -ErrorAction "SilentlyContinue" -InputObject $obj -MemberType NoteProperty  -Name (($prop.name -split "\." )[-1]) -Value  $row."$($prop.name)"
            }                       
            #Add aliases 
            foreach ($prop in ($PropertyAliases.Keys | where-object {  ($row."$($propertyAliases.$_)" -isnot [system.dbnull] ) -and
                                                                       ($row."$($propertyAliases.$_)" -ne $null )})) {
                Add-member -ErrorAction "SilentlyContinue" -InputObject $obj -MemberType AliasProperty -Name $prop -Value ($propertyAliases.$prop  -split "\." )[-1] 
            }
            #Overwrite duration as a timespan not as 100ns ticks
            If ($obj.duration) { $obj.duration =([timespan]::FromMilliseconds($obj.Duration / 10000) )}
            $obj
        }
    }                               
 }}
} 
function New-Overlay {
    <#
        .Synopsis
            Creates a new transparent JPG containing text to use as an overlay
        .Description
            Creates a new transparent JPG containing text to use as an overlay
        .Example
            C:\PS>$Overlay = New-overlay -text "© James O'Neill 2009" -size 32 -TypeFace "Arial"  -color red -filename "$Pwd\overLay.jpg"
            C:\PS>$filter = Add-OverlayFilter $overlay -passthru
            
            
            Creates an overlay adds it to the filter list
        .Parameter Text
            The text you want to set in the overlay
        .Parameter Size
            The font size (default 32)
        .Parameter TypeFace
            The font face (Default Arial)
        .Parameter Color
            The font Color (Default grey; type [system.drawing.color]:: and then tab through to see the more obscure color names available)
        .Parameter Filename
            The name for the overlay file (default "Overlay.jpg" in the current folder)
    #>
param ([string]$text="All rights reserved by the copyright owner", [int]$size = 32, [string]$TypeFace="Arial", [system.drawing.color]$color=[system.drawing.color]::Gray, [string]$filename="$Pwd\overLay.jpg" )
    $width= $text.length * $size   # This is a close enough approximation - 
    $height = $size * 1.5
    $Thumbnail = new-object System.Drawing.Bitmap($width, $height, [System.Drawing.Imaging.PixelFormat]::Format16bppRgb565)           
    $graphic=[system.drawing.graphics]::FromImage($thumbnail)                                                                         
    if ($color -eq [system.drawing.color]::Black) {$b=new-object system.drawing.solidbrush([system.drawing.color]::white)
                                                    $graphic.FillRectangle($b,0,0,$width,$height)                                           
    } 
    $font=new-object system.drawing.font($typeFace,$size)                                                                                                             
    $brush=new-object system.drawing.solidbrush($color)                                                      
    $graphic.DrawString($text,$font ,$brush, 0 , 0 )         
    if ($color -eq [system.drawing.color]::Black) {$Thumbnail.MakeTransparent([system.drawing.color]::White)} else {$Thumbnail.MakeTransparent([system.drawing.color]::Black)}                                                                       
    $Thumbnail.Save($filename)
    Get-Image $filename
}
function Save-image {
   <#
        .Synopsis
            Saves a Windows Image Acquisition image
        .Description
            Saves a Windows Image Acquisition image
        .Example   
            C:\ps> $image | Save-image -NoClobber -fileName {$_.FullName -replace ".jpg$","-small.jpg$"}
            Saves the JPG image(s) in $image, in the same folder as the source JPG(s), appending
            -small to the file name(s), so that "MyImage.JPG" becomes "MyImage-Small.JPG"
            Existing images will not be overwritten.
        .Parameter image
            The image or images the filter will be applied to; images may be passed via the pipeline.
            If multiple images are passed either no filename must be included 
            (so the image will be saved under its original name), or the fileName must be a code block,
            otherwise the images will all be written over the same file. 
        .Parameter passThru
            If set, the image or images will be emitted onto the pipeline       
        .Parameter filename
            If not set the existing file will be overwritten. The filename may be a string, 
            or be a script block - as in the example
        .Parameter NoClobber
            specifies the target file should not be over written if it already exists
    #>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
      [Parameter(ValueFromPipeline=$true,Mandatory=$true)]
      $image,
      [parameter()][Alias("Path","FullName")][ValidateNotNullOrEmpty()]
      $fileName = $image.fullName,
      [switch]$passThru,
      [switch]$NoClobber, $psc )

process {
      if ( $psc -eq $null )            { $psc = $pscmdlet }   ; if (-not $PSBoundParameters.psc) {$PSBoundParameters.add("psc",$psc)}
      if ( $image.count -gt 1       )  { [Void]$PSBoundParameters.Remove("Image") ;  $image | ForEach-object {Save-Image -Image $_ @PSBoundParameters }  ; return}
      if ($filename -is [scriptblock]) {$fname = Invoke-Expression $(".{$filename}") }
      else                             {$fname = $filename } 
      if (test-path $fname)            {if     ($noclobber) {write-warning "$fName exists and WILL NOT be overwritten"; if ($passthru) {$image} ; Return }
                                        elseIF ($psc.shouldProcess($FName,"Delete file")) {Remove-Item  -Path $fname -Force -Confirm:$false }
                                        else   {Return}
      }  
      if ((Test-Path -Path $Fname -IsValid) -and ($pscmdlet.shouldProcess($FName,"Write image")))  { $image.SaveFile($FName) }
      if ($passthru) {$image} 
   }
}
