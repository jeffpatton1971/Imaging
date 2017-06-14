Function Get-PentaxMakerNoteProperty {
<#
        .Synopsis
            Returns an single item of Pentax specific maker note information
        .Description
            Returns an single item of Pentax specific maker note information
        .Example
            C:\PS> Get-PentaxMakerNoteProperty -image $image -exifID 5 
            Returns the Camera model ID code. 
        .Parameter image
            The image from which the data will be read 
        .Parameter ExifID
            The ID of the required data field
    #>
Param ([int]$exifID,   [__ComObject] $image)
       $MakerID ="AOC" 
       try   { [byte[]]$wholeMakerNote= $image.Properties.Item("37500").value.binaryData }
       Catch { Write-Warning "Error getting maker note - proably doesn't exist" ; return   }
       0..$MakerId.length | ForEach-Object {if ($wholeMakerNote[$_] -ne [byte]($makerid[$_]) ) {write-debug "Wrong maker ID"; Return}}
       $PrevFieldCode = 0
       for ($i = 8 ; $i -lt $WholeMakerNote.count -1 ; $i += 12 ) {
            $FieldCode = 256 * $WholeMakerNote[$i] + $WholeMakerNote[$i + 1]
            If (($FieldCode -lt $PrevFieldCode) -or ($fieldcode -gt $exifID)) {write-debug "Field code not found"; return} else {$PrevFieldCode = $FieldCode}
            #If (($WholeMakerNote[$i + 4] -ne 0) -Or ($WholeMakerNote[$i + 5] -ne 0)) {return } # don't recall why
            If ($FieldCode -eq $ExifID) {#& write-host $wholemakernote[$i+3] 
                if  ($WholeMakerNote[$i + 3] -eq 3) {$S = [string](256 * $WholeMakerNote[$i + 8] + $WholeMakerNote[$i + 9])
                         If ($WholeMakerNote[$i + 7] -eq 2) {$s = $s + " , " + [String](256 * $WholeMakerNote[$i + 10] + $WholeMakerNote[$i + 11]) }
                         Return $s}
                if (($WholeMakerNote[$i + 3] -eq 4) -and ($WholeMakerNote[$i + 7] -eq 1)) { Return [String](16777216 * $WholeMakerNote[$i + 8] + 65536 * $WholeMakerNote[$i + 9] + 256 * $WholeMakerNote[$i + 10] + $WholeMakerNote[$i + 11])}
                if (( @(1, 6, 7) -contains $WholeMakerNote[$i + 3]) -and ($WholeMakerNote[$i + 7] -le 4))  {($i +8)..($i +7 + $WholeMakerNote[$i + 7] ) | foreach -Begin {$s=""} -process {$s = $s + [string]$WholeMakerNote[$_]  + " "  } -end {return $s} }
            }
        }         
}

function Get-PentaxExif {
    <#
        .Synopsis
            Returns an object containing Pentax specific maker note information
        .Description
            Returns an object containing Pentax specific maker note information
        .Example
            Get-pentaxExif $image
        .Parameter image
            The image from which the data will be read 
        .Parameter full
            Pentax data only or all adata 

    #>
Param ([Parameter(ValueFromPipeline=$true,Mandatory=$true)]$image , [Switch]$full)
    if ($image -is [system.io.fileinfo] ) {$image = $image.FullName }
    if ($image -is [String]             ) {$image = Get-Image $image}    
    if ($image.count -gt 1              ) {$image | ForEach-Object {Get-PentaxExif $_ -Full:$full} ; Return}
    $s = Get-PentaxMakerNoteProperty -image $image -exifID 5 
    if ($s -eq $null) {Write-Warning "No camera ID was found, giving up"; return}
    Switch ($s) {
           "13"{$PentaxModel = "Optio 430"}
        "76450"{$PentaxModel = "*ist-D"   }
        "76830"{$PentaxModel = "K10D"     }
        "77240"{$PentaxModel = "K7"       }
        Default{$PentaxModel = "Unknown"  }
    }
    
     
           $s           = (Get-PentaxMakerNoteProperty -image $image -exifID 93).Split(" ") 
           $d           = (Get-PentaxMakerNoteProperty -image $image -exifID 6).Split(" ") 
           $t           = (Get-PentaxMakerNoteProperty -image $image -exifID 7).Split(" ")
    [int64]$dno         =               (16777216 * $d[0]) + (65536 * $d[1]) + (256 * $d[2]) +($d[3])                                              
    [int64]$tno         =               (16777216 * $t[0]) + (65536 * $t[1]) + (256 * $t[2]) +($t[3])                                              
    [int64]$sno         = 4294967296 -  (16777216 * $s[0]) - (65536 * $s[1]) - (256 * $s[2]) -($s[3]) - 1                                         
    $PentaxShutterCount =  $Sno -bXor $dNo -bXor $tNo                

    Switch (Get-PentaxMakerNoteProperty -image $image -exifID 31){
        "0"{$PentaxSaturation = "Low"      }
        "1"{$PentaxSaturation = "Normal"   }
        "2"{$PentaxSaturation = "High"     }
        "3"{$PentaxSaturation = "Med-Low"  }
        "4"{$PentaxSaturation = "Med-High" }
        "5"{$PentaxSaturation = "Very Low" }
        "6"{$PentaxSaturation = "Very High"}
    Default{$PentaxSaturation = "Unknown"  }
    }

    Switch (Get-PentaxMakerNoteProperty -image $image -exifID 32){
        "0"{$PentaxContrast = "Low"       }
        "1"{$PentaxContrast = "Normal"    }
        "2"{$PentaxContrast = "High"      } 
        "3"{$PentaxContrast = "Med-Low"   }
        "4"{$PentaxContrast = "Med-High"  }
        "5"{$PentaxContrast = "Very Low"  }
        "6"{$PentaxContrast = "Very High" }
    Default{$PentaxContrast = "Unknown"   }
    }

    Switch (Get-PentaxMakerNoteProperty -image $image -exifID 33){
        "0"{$PentaxSharpening = "Low"      }
        "1"{$PentaxSharpening = "Normal"   }
        "2"{$PentaxSharpening = "High"     }
        "3"{$PentaxSharpening = "Med-Low"  }
        "4"{$PentaxSharpening = "Med-High" }
        "5"{$PentaxSharpening = "Very Low" }
        "6"{$PentaxSharpening = "Very High"}
        Default{$PentaxSharpening = ""     }
    }

    Switch (Get-PentaxMakerNoteProperty -image $image -exifID 79) {
        "0"{$PentaxImageTone = "Natural"   }
        "1"{$PentaxImageTone = "Bright"    }
        "2"{$PentaxImageTone = "Portrait"  }
        "3"{$PentaxImageTone = "Landscape" }
        "4"{$PentaxImageTone = "Vibrant"   }
        "5"{$PentaxImageTone = "Monochrome"}
    Default{$PentaxImageTone = "Unknown"   }
    }

    $s= (Get-PentaxMakerNoteProperty -image $image -exifID 92).Split(" ")
    $t="" 
    If ([int]$s[1] –band 1)  {$t =      "SR Enabled."  } else {$t =     "SR Disabled."    }
    If ([int]$s[0] –band 1)  {$t = $t + " Stabilized." } else {$t = $t +" Not Stabilized."}
    If ([int]$s[0] –band 64) {$t = $t + " Not Ready."   }
    If ([int]$s[3] -band 1)  {$t = $t + " SR Focal Length: $([int]$S[3]*4)"} else {$t = $t +  " SR Focal Length: $([int]$S[3]/2)mm"}
    $PentaxShakeReduction =   $t

    $s = (Get-PentaxMakerNoteProperty -image $image -exifID 51).Split(" ")
    $t=""
    Switch ($S[0]){
        "0"{$t  = "Program"                }
        "1"{$t  = "Shutter Speed Priority" }
        "2"{$t  = "Program AE"             }
        "3"{$t  = "Manual"                 }
        "5"{$t  = "Portrait"               } 
        "6"{$t  = "Landscape"              }
        "8"{$t  = "Sport"                  }
        "9"{$t  = "Night Scene"            }
        "11"{$t = "Soft"                   }
        "12"{$t = "Surf & Snow"            }
        "13"{$t = "Candlelight"            }
        "14"{$t = "Autumn"                 }
        "15"{$t = "Macro"                  }
        "17"{$t = "Fireworks"              }
        "18"{$t = "Text"                   }
        "19"{$t = "Panorama"               }
        "30"{$t = "Self-Portrait"          }
        "31"{$t = "Illustrations"          }
        "33"{$t = "Digital Filter"         }
        "37"{$t = "Museum"                 }
        "38"{$t = "Food"                   }
        "40"{$t = "Green Mode"             }
        "49"{$t = "Light Pet"              }
        "50"{$t = "Dark Pet"               }
        "51"{$t = "Medium Pet"             }
        "53"{$t = "Underwater"             }
        "54"{$t = "Candlelight"            }
        "55"{$t = "Natural Skin Tone"      }
        "56"{$t = "Synchro Sound Record"   }
        "58"{$t = "Frame Composite"        }
        "60"{$t = "Kids"                   }
        "61"{$t = "Blur Reduction"         }
        "255"{$t= "Digital Filter"        }
    }
    If ($S[2] -eq "0") {$PentaxPictureMode = $t + ": 1/2 EV Steps"} else {$PentaxPictureMode = $t + ": 1/3 EV Steps"}

    $s= (Get-PentaxMakerNoteProperty -image $image -exifID 52).Split(" ")
    $t = ""
    Switch ($S[0]){
        "0"{$t = "Single-frame. " }
        "1"{$t = "Continuous. "}
        "2"{$t = "Continuous [Hi]. "}
    }
    Switch ($S[1]){
        "0"{$t = $t + "No timer. " }
        "1"{$t = $t + "Self-timer, 12 sec. " }
        "2"{$t = $t + "Self-timer, 2 sec. " }
    }
    Switch ($S[2]){
        "0"{$t = $t + "Shutter button. "}
        "1"{$t = $t + "Remote control,3 sec delay. "}
        "2"{$t = $t + "Remote control."}
    }
    If ($S[3] = "0") {$PentaxDriveMode = $t + " Single Exposure."} else {$PentaxDriveMode = $t + " Multiple Exposure."}

    # ? Add 0x0032 imageProcessing , and 0x0041; 0x0048 AELock ; 0x0049 Noise Reduction?  

    Switch  (Get-PentaxMakerNoteProperty -image $image -exifID 13) {
        "0" {$PentaxFocusMode = "Normal"   }
        "1" {$PentaxFocusMode = "Macro"    }
        "2" {$PentaxFocusMode = "Infinity" }
        "3" {$PentaxFocusMode = "Manual"   }
        "5" {$PentaxFocusMode = "Pan-Focus"}
        "16"{$PentaxFocusMode = "AF-S"     }
        "17"{$PentaxFocusMode = "AF-C"     }
     Default{$PentaxFocusMode = "Unknown"  }
    }

    Switch  (Get-PentaxMakerNoteProperty -image $image -exifID 14) {
        "1"{$PentaxFocusPoint = "Top Left"        }
        "2"{$PentaxFocusPoint = "Top Center"      }
        "3"{$PentaxFocusPoint = "Top Right"       }
        "4"{$PentaxFocusPoint = "Middle Far-Left" }
        "5"{$PentaxFocusPoint = "Middle Left"     }
        "6"{$PentaxFocusPoint = "Middle Center"   }
        "7"{$PentaxFocusPoint = "Middle Right"    }
        "8"{$PentaxFocusPoint = "Middle Far-Right"}
        "9"{$PentaxFocusPoint = "Bottom Left"     }
       "10"{$PentaxFocusPoint = "Bottom Center"   }
       "11"{$PentaxFocusPoint = "Bottom Right"    }
    "65534"{$PentaxFocusPoint = "Fixed Center"    }
    "65535"{$PentaxFocusPoint = "Auto"            }
    Default{$PentaxFocusPoint = "Unknown"         }
    }

    $s =(Get-PentaxMakerNoteProperty -image $image -exifID 63).Split(" ")
    $i = (256 * [int]$S[0]) + [int]$s[1]
    Switch ($i) { 
         256{$PentaxLens = "M-Series"                               }
         512{$PentaxLens = "A-Series"                               }
         812{$PentaxLens = "Sigma 10-20 F4-5.6 EX DC"               }
         814{$PentaxLens = "Sigma 100-300 F4.5-6.7"                 }
        1039{$PentaxLens = "SMC PENTAX-FA 28-105mm F4-5.6 [IF]"     }
        1048{$PentaxLens = "SMC-PENTAX-FA 77mm F1.8 Limited"        }
        1071{$PentaxLens = "SMC PENTAX-FA J 18-35mm F4-5.6 AL"      }
        1284{$PentaxLens = "SMC PENTAX-FA 50mm F1.4"                }
        2023{$PentaxLens = "SMC PENTAX-DA 18-250 F3.5-63 ED EL[IF]" }
     Default{$PentaxLens = "ID= $i"                                 }
    }
    Switch (Get-PentaxMakerNoteProperty -image $image -exifID 23) {
        "0"{$PentaxMetermode = "Multi-Segment"   }
        "1"{$PentaxMetermode = "Center-Weighted" }
        "2"{$PentaxMetermode = "Spot"            }
    Default{$PentaxMetermode = "Unknown"         }
    }

    Switch (Get-PentaxMakerNoteProperty -image $image -exifID 8) {
        "0"{$PentaxQuality = "Good"    }
        "1"{$PentaxQuality = "Better"  }
        "2"{$PentaxQuality = "Best"    }
        "3"{$PentaxQuality = "TIFF"    }
        "4"{$PentaxQuality = "RAW"     }
    Default{$PentaxQuality = "Unknown" }
    }

    $PentaxTemperature = (Get-PentaxMakerNoteProperty -image $image -exifID 71)

    $h= @{}
    if ($full) {$e= get-exif $image ; get-member -input $e -MemberType noteproperty | % {$h.add($_.name, $e."$($_.name)") }    }

    New-Object PSObject -Property (@{PentaxContrast       =  $PentaxContrast   
                                     PentaxShutterCount   =  $PentaxShutterCount
                                     PentaxSaturation     =  $PentaxSaturation  
                                     PentaxSharpening     =  $PentaxSharpening  
                                     PentaxImageTone      =  $PentaxImageTone
                                     PentaxShakeReduction =  $PentaxShakeReduction
                                     PentaxPictureMode    =  $PentaxPictureMode  
                                     PentaxDriveMode      =  $PentaxDriveMode
                                     PentaxFocusMode      =  $PentaxFocusMode   
                                     PentaxFocusPoint     =  $PentaxFocusPoint
                                     PentaxLens           =  $PentaxLens 
                                     PentaxMetermode      =  $PentaxMetermode      
                                     PentaxModel          =  $PentaxModel 
                                     PentaxQuality        =  $PentaxQuality
                                     PentaxTemperature    = "$PentaxTemperature°c" } +$h ) 
}