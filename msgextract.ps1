#msgextract.ps1
#Parses Metadata of .msg files
#Version 1
#James Saunders
#07/24/2013


#date
$date = Get-Date -format D

#Set directories
$BaseDir = "D:\_IT\SPAM\msgextract"

$inputDir = (Join-Path $BaseDir "in")
$outputDir = (Join-Path $BaseDir "out")
$tmpDir = (Join-Path $BaseDir "tmp")
$timestamp = $(get-date -f "yyyy-MM-dd HH-mm-ss")

$webReportFile = (Join-Path $outputDir "$timestamp-webreport.html")

Function meta_extract($_msg){
    #Write-Output($_msg.Name)
    
    $blah = (Join-Path $tmpDir $date) #create folder with date
    $blah = (Join-Path $blah $_msg.Name) # create subfolder with name of msg
    #$7z x $_msg -o"$OutputDir" |out-null 
    & "C:\Program Files\7-Zip\7z.exe" "x" $_msg "-o$blah" | out-null #unpack msgs
}

Function fileSorter{
    $msg_filters = 
    @(
        #@{ptrn='001A'; class="Message class"},
        #@{ptrn='0037'; class="Subject"},
        #@{ptrn='0040'; class="Received by name"},
        #@{ptrn='0042'; class="Sent repr name"},
        #@{ptrn='0044'; class="Rcvd repr name"},
        #@{ptrn='004D'; class="Org author name"},
        #@{ptrn='0050'; class="Reply rcipnt names"},
        #@{ptrn='005A'; class="Org sender name"},
        #@{ptrn='0064'; class="Sent repr adrclass"},
        #@{ptrn='0065'; class="Sent repr email"},
        #@{ptrn='0070'; class="Topic"},
        #@{ptrn='0075'; class="Rcvd by adrclass"},
        #@{ptrn='0076'; class="Rcvd by email"},
        #@{ptrn='0077'; class="Repr adrclass"},
        #@{ptrn='0078'; class="Repr email"},
        @{ptrn='007d'; class="Message header"}#,
        #@{ptrn='0C1A'; class="Sender name"},
        #@{ptrn='0C1E'; class="Sender adr class"},
        #@{ptrn='0C1F'; class="Sender email"},
        #@{ptrn='0E02'; class="Display BCC"},
        #@{ptrn='0E03'; class="Display CC"},
        #@{ptrn='0E04'; class="Display To"},
        #@{ptrn='1000'; class="Message body"},
        #@{ptrn='1046'; class="Sender email"},
        #@{ptrn='3001'; class="Display name"},
        #@{ptrn='3002'; class="Address class"},
        #@{ptrn='3003'; class="Email address"}
    )

    $data = gci $tmpDir\*\*\*.*
    $filename="Message header.txt" 
    $outName = (Join-Path $outputDir "$timestamp-$filename")
    foreach($f in $data){       
      
      foreach($c in $msg_filters){                                    
        if ($f -match $c.ptrn){ 
          #Filename from Filtermatch
          #$filename=$c.class+".txt" 
          #combine the path                               
          #$outName = (Join-Path $outputDir "$timestamp-$filename")                  

          ac $outName (gc $f)
          #$msgdata = gc $f|out-file $outName
          

          #write to web report
          # $tmphtml = "<div><msg>" + $c.class + "</msg>"
          #ac $webReportFile $tmphtml
          #$tmpdata = gc $outName
          # $tmphtml += "<p>" + $f + "</p></div>"
          # ac $webReportFile $tmphtml
        }
      }
    }
    (gc $outName -readcount 0) -replace "`0", "" | sc $outName
}

Function Process-ReceivedFrom 
{ 
  Param($text) 
  $regexFrom1 = 'Received: from([\s\S]*?)by([\s\S]*?)with([\s\S]*?);([(\s\S)*]{32,36})(?:\s\S*?)' 
  $fromMatches = $text | Select-String -Pattern $regexFrom1 -AllMatches 
  if ($fromMatches) 
  { 
          $rfArray = @() 
          $fromMatches.Matches | foreach{ 
          $from = Clean-string $_.groups[1].value 
          $by = Clean-string $_.groups[2].value 
          $with = Clean-string $_.groups[3].value 
              Switch -wildcard ($with) 
              { 
               "SMTP*" {$with = "SMTP"} 
               "ESMTP*" {$with = "ESMTP"} 
               default{} 
              } 
          $time = Clean-string $_.groups[4].value 
          $fromhash = @{ 
              ReceivedFromFrom = $from 
              ReceivedFromBy = $by 
              ReceivedFromWith = $with 
              ReceivedFromTime = [Datetime]$time 
          }         
          $fromArray = New-Object -TypeName PSObject -Property $fromhash         
          $rfArray += $fromArray         
          } 
          $rfArray 
  } 
  else 
  { 
      return $null 
  } 
} 

Function Process-ReceivedBy 
{ 
  Param($text) 
  $regexBy1 = 'Received: by ' 
  $regexBy2 = 'Received: by ([\s\S]*?)with([\s\S]*?);([(\s\S)*]{32,36})(?:\s\S*?)' 
  $regexBy3 = 'Received: by ([\s\S]*?);([(\s\S)*]{32,36})(?:\s\S*?)' 
  $byMatches = $text | Select-String -Pattern $regexBy1 -AllMatches 
   
  if ($byMatches) 
  { 
      $byMatches = $text | Select-String -Pattern $regexBy2 -AllMatches 
      if($byMatches) 
      { 
          $rbArray = @() 
          $byMatches.Matches | foreach{ 
          $by = Clean-string $_.groups[1].value 
          $with = Clean-string $_.groups[2].value 
              Switch -wildcard ($with) 
              { 
               "SMTP*" {$with = "SMTP"} 
               "ESMTP*" {$with = "ESMTP"} 
               default{} 
              } 
          $time = Clean-string $_.groups[3].value 
          $byhash = @{ 
              ReceivedByBy = $by 
              ReceivedByWith = $with 
              ReceivedByTime = [Datetime]$time 
          }         
          $byArray = New-Object -TypeName PSObject -Property $byhash         
          $rbArray += $byArray         
          } 
          $rbArray 
      } 
      else 
      { 
          $rbArray = @() 
          $byMatches = $text | Select-String -Pattern $regexBy3 -AllMatches 
          $byMatches.Matches | foreach{ 
          $by = Clean-string $_.groups[1].value 
          $with = "" 
          $time = Clean-string $_.groups[2].value 
          $byhash = @{ 
              ReceivedByBy = $by 
              ReceivedByWith = $with 
              ReceivedByTime = [Datetime]$time 
          } 
          $byArray = New-Object -TypeName PSObject -Property $byhash         
          $rbArray += $byArray         
          } 
          $rbArray 
      } 
  } 
  else 
  { 
      return $null 
  } 
} 

Function Clean-String 
{ 
  Param([string]$inputString)   
   $inputString = $inputString.Trim() 
   $inputString = $inputString.Replace("`r`n","")   
   $inputString = $inputString.Replace("`t"," ")  
   $inputString 
} 

Function attachmentSorter{
    $attachment_filters = 
                    @(
                                    @{ptrn='3701'; class="Attachment data"},
                                    @{ptrn='3703'; class="Attach extension"},
                                    @{ptrn='3704'; class="Attach filename"},
                                    @{ptrn='3707'; class="Attach long filenm"},
                                    @{ptrn='370E'; class="Attach mime tag"}
                    )
    
    $data = gci $ExtractedMSGDir
    foreach($f in $data){
                    foreach($c in $attachment_filters){
                                    if ($f -match $c.ptrn){
                                                    $OutName = $ExtractedMSGDir + $c.class + ".txt"
                                                    $msgdata = gc $ExtractedMSGDir+$f|out-file $outName
                                                    (gc $outName) -replace "`0", "" | sc $outName
                                                    $tmphtml = "<div><msg>" + $c.class + "</msg>"
                                                    ac $ExtractedMSGDir"webreport.html" $tmphtml
                                                    $tmpdata = gc $outName
                                                    $tmphtml = "<p>" + $tmpdata + "</p></div>"
                                                    ac $ExtractedMSGDir"webreport.html" $tmphtml
                                    }
                    }
    }
                                
}
Function Process-FromByObject 
{ 
  Param([PSObject[]]$fromObjects,[PSObject[]]$byObjects) 
  [int]$hop=0 
  $delay="" 
  $receivedfrom=$receivedby=$receivedtime=$receivedwith=$null 
  $prevTime=$null 
  $time=$null 
  $finalArray = @() 
      if($byObjects) 
      {         
       $byObjects = $byObjects[($byObjects.Length-1)..0] # Reversing the Array 
       for($index = 0;$index -lt $byobjects.Count;$index++) 
          { 
              if($index -eq 0) 
              { 
                  $hop=1 
                  $delay="*" 
                  $receivedfrom = "" 
                  $receivedby = $byobjects[$index].ReceivedByBy 
                  $with = $byobjects[$index].ReceivedByWith 
                  $time = $byobjects[$index].ReceivedBytime 
                  $time = $time.touniversaltime() 
                  $prevTime = $time 
                  $finalHash = @{ 
                      Hop   = $hop 
                      Delay = $delay 
                      From  = $receivedfrom 
                      By       = $receivedby 
                      With  = $with 
                      Time  = $time 
                      }                 
                  $obj = New-Object -TypeName PSObject -Property $finalHash 
                  $finalArray += $obj                 
              } 
              else 
              { 
                  $hop = $index+1                 
                  $receivedfrom = "" 
                  $receivedby = $byobjects[$index].ReceivedByBy 
                  $with = $byobjects[$index].ReceivedByWith 
                  $time = $byobjects[$index].ReceivedBytime 
                  $time = $time.touniversaltime()                 
                  $delay = $time - $prevTime 
                  $delay = $delay.totalseconds 
                  if ($delay -le -1) {$delay = 0}                 
                  $prevTime = $time 
                                  $finalHash = @{ 
                      Hop   = $hop 
                      Delay = $delay 
                      From  = $receivedfrom 
                      By       = $receivedby 
                      With  = $with 
                      Time  = $time 
                      }                 
                  $obj = New-Object -TypeName PSObject -Property $finalHash 
                  $finalArray += $obj 
              } 
          } 
       $lastHop = $hop 
        
      } 
      $hop = $lastHop 
      if($fromObjects) 
      {         
       $fromObjects = $fromObjects[($fromObjects.Length-1)..0] #Reversing the Array 
       for($index = 0;$index -lt $fromobjects.Count;$index++) 
          {             
           
                  $hop = $hop + 1 
                  $receivedfrom = $fromobjects[$index].ReceivedFromFrom 
                  $receivedby = $fromobjects[$index].ReceivedFromBy 
                  $with = $fromobjects[$index].ReceivedFromWith 
                  $time = $fromobjects[$index].ReceivedFromTime 
                  $time = $time.touniversaltime()                 
                  if($prevTime) 
                  { 
                      $delay = $time - $prevTime 
                      $delay = $delay.totalseconds 
                  } 
                  else 
                  { 
                      $delay = "*" 
                  }                 
                  $prevTime = $time 
                  $finalHash = @{ 
                      Hop   = $hop 
                      Delay = $delay 
                      From  = $receivedfrom 
                      By       = $receivedby 
                      With  = $with 
                      Time  = $time 
                      }                 
                  $obj = New-Object -TypeName PSObject -Property $finalHash 
                  $finalArray += $obj 
               
          } 
        
      } 
  $finalArray 
} 
Function Main{
  #ac $webReportFile "<html><head><style>msg { color:black; font-size:20px; background: #7abcff;  font-family:Verdana, Helvetica, Arial, sans-serif;  } h1{font-family:Verdana, Helvetica, Arial, sans-serif;} div{border:1px solid #333; border-width:1px 0;} </style><title>Email Metadata Report | $(get-date -f yyyy-MM-dd)</title></head><body><h1>Email Metadata Report | $(get-date -f yyyy-MM-dd)</h1>"

  $files = Get-ChildItem $inputDir\*.msg
  $files|foreach{
    Write-host $_
    #meta_extract $_ 
    Write-host $_.Name.replace(".msg","").replace("FW  ","").replace("Fw  ","") -foregroundcolor black -backgroundcolor yellow
    Write-host "Metadata Extracted!" -foregroundcolor red -backgroundcolor yellow
    #fileSorter $_
  }
  #ac $webReportFile "</body></html>"
  #ii $outputDir"webreport.html"
  #ri $tmpDir"\*.*" -recurse
}

Function Process 
{ 
  $filename="Message header.txt"
  $InputFileName = (Join-Path $outputDir "$timestamp-$filename")
  $text = [System.IO.File]::OpenText($InputFileName).ReadToEnd() 
  $fromObject = Process-ReceivedFrom -text $text 
  # $byObject = Process-ReceivedBy -text $text 
   
  $finalArray = Process-FromByObject $fromObject $byObject 
  #Write-Output $finalArray 
  $finalArray | Export-Csv "$outputDir\$timestamp-headerdata.csv"
} 

Main
#Process

######CL######
#7/24 V1 created
#####Future####
#Loop thru attachments
