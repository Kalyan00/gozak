$glob = @{
   dir = "c:\GosZakup.DB\"
   Days = 20
   MaxDownloadedErrorCounter = 10


   OKPD2_Remove = @(
      #"10.", #Продукты пищевые
      "80.10.12", #Услуги охраны
      "86",  #Услуги в области здравоохранения
      "")
   minPrice = 80000
}

function FtpStore {"$($glob.dir)FTP\"}
function AttachStore {"$($glob.dir)attachments\"}


function Init {
   Add-Type -AssemblyName System.IO.Compression
   Add-Type -AssemblyName System.IO.Compression.FileSystem

   Check-Proxy
   cdir $glob.dir |out-null
}

function cdir($Path) {
   if(-not (Test-Path $Path))
   {
      mkdir $Path | out-null
   }
   $Path
}

function Check-Proxy(){

   $settingproxy = cat "$($glob.dir)proxy.txt"|?{-not $_.StartsWith("#")}| select -first 1
   
   if($settingproxy)
   {
      [System.Net.WebRequest]::DefaultWebProxy = New-Object System.Net.WebProxy($settingproxy)
   }
}

function DictToObj{
   Process{
   $result = new-object PSObject
   $_.GetEnumerator()|%{
      $result | Add-Member -Name $_.Name -Value $_.Value -MemberType Noteproperty
   }

   $result

   }
}

function String-Join($divider, $array){
   if(!$array) {return ""}
   [array]$array = $array|?{$_}
   if(!$array) {return ""}

   [String]::Join($divider,$array)
}

function StreamToString($stream) {
   $readStream = new-object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
   $readStream.ReadToEnd()
}

function GetFtpStream($filename, $lambda) {

   $url = "ftp://ftp.zakupki.gov.ru/$filename"
   [System.Net.FtpWebRequest]$WR = [System.Net.WebRequest]::Create($url)
   $WR.Credentials = New-Object System.Net.NetworkCredential("free","free")
   if($filename.endswith('/'))
   {
   $WR.Method=[System.Net.WebRequestMethods+Ftp]::ListDirectory
   }
   $WRStream = $WR.GetResponse()
   $stream = $WRStream.GetResponseStream()
   $lambda.Invoke($stream)
   $stream.Close()
   $WRStream.Close()
   $WR.Abort()
}

function GetFiles($dir){

   $fileStr = GetFtpStream $dir {param($stream) StreamToString $stream}
   $files = $fileStr.Split("`n`r") |?{$_ -match $glob.DaysMatch} |%{$Matches[0]}
   $files |%{"$dir$_"}
}


function GetAllFiles{
   $glob.DaysMatch = "(" + (String-Join "|" (0..$glob.Days | % {[DateTime]::Now.AddDays(-$_).ToString("yyyyMMdd")}))+")"
   $glob.DaysMatch = "(notification_[^`"]*?$($glob.DaysMatch)\d\d_\d\d\d.xml.zip)"

   GetFiles "/fcs_regions/Moskva/notifications/currMonth/"
   GetFiles "/fcs_regions/Moskovskaja_obl/notifications/currMonth/"
   if ($glob.Days -gt [DateTime]::Now.Day) {
      GetFiles "/fcs_regions/Moskva/notifications/prevMonth/"
      GetFiles "/fcs_regions/Moskovskaja_obl/notifications/prevMonth/"

   }
}

function SaveFtpZips {
   $store = cdir(FtpStore)

   GetAllFiles |%{
      $remote = $_
      $fileName = $_.Split('/')|select -last 1
      $local = "$store$filename"

      if(! (Test-Path $local))
      {
         GetFtpStream $remote {
            param($stream)
            $outputStream = new-object IO.FileStream($local, [IO.FileMode]::Create)
            $stream.CopyTo($outputStream)
            $outputStream.Close()
         }
      }
   }
}

function ClearFtpZips {
   $store = cdir(FtpStore)

   dir "$($store)*.*"|?{! ($_.name -match $glob.DaysMatch)}|%{
      del $_.FullName
   }
}

function filterZK44($ZK44_Object) {
   if($ZK44_Object.maxPrice -le $glob.minPrice) { return $false }

   if($ZK44_Object.dateOpen  -le [DateTime]::Now) { return $false }

   if($ZK44_Object.purchaseObjects | ?{$_.OKPD2Code}| ?{
      $OKPD2Code = $_.OKPD2Code;
      $glob.OKPD2_Remove|?{$_}|?{ $OKPD2Code.StartsWith($_)}
      })
   {
      return $false
   }

   $true
}

function ReadZips() {
   if(!$glob.ZK44_Objects)
   {
      write-host("   ReadZips start")
      $glob.ZK44_Objects = @{}

      dir "$(FtpStore)*.zip" |%{
         $zipFile = $_.Name
         $zip = [System.IO.Compression.ZipFile]::Open($_.FullName,"Read")
         $entries = $zip.Entries |?{$_.FullName -match "ZK44_.*.xml"}

         $entries|%{

            $reqFile = $_.FullName
            [xml]$xml = StreamToString $_.Open()

            $notify = ReadZips_internal $xml $zipFile $reqFile
            if(!$glob.ZK44_Objects[$notify.regNumber])
            {
               $glob.ZK44_Objects[$notify.regNumber] = @{}
            }
            $glob.ZK44_Objects[$notify.regNumber][$notify.notifyId] = $notify
         }
      }

      $glob.ZK44_Objects.values |%{
         $max = ($_.keys | measure-object  -maximum).maximum
         $_.GetEnumerator() |?{$_.key -ne $max}|%{
            $_.value.is_Good = $false
            $_.value.is_Double = $true
         }
      }
      write-host("   ReadZips end")
   }

   $glob.ZK44_Objects.values|%{$_.values}
}

function ReadZips_internal($xml,$zipFile,$reqFile) {
   [array]$purchaseObjects = $xml.DocumentElement.fcsNotificationZK.lot.purchaseObjects.purchaseObject|%{
      @{
         OKPD2 = $_.OKPD2.name
         OKPD2Code = $_.OKPD2.code
         name = $_.name
         edism = $_.OKEI.nationalCode
         quantity = $_.quantity.value
         priceStr = ([double]$_.price).ToString('C')
         sum = $_.sum
         sumStr = ([double]$_.sum).ToString('C')
      }|DictToObj
   }
   [array]$attachments = $xml.DocumentElement.fcsNotificationZK.attachments.attachment |%{
      @{
         fileName = $_.fileName
         fileSize = $_.fileSize
         docDescription = $_.docDescription
         url = $_.url
         # url = "https://az818661.vo.msecnd.net/providers/providers.masterList.feed.swidtag"
      }|DictToObj
   }
   $dateOpen = [DateTime]$xml.DocumentElement.fcsNotificationZK.procedureInfo.opening.date

   $result = @{
      zipFile = $zipFile
      reqFile = $reqFile
      regNumber = $xml.DocumentElement.fcsNotificationZK.purchaseNumber
      notifyId = [int]$xml.DocumentElement.fcsNotificationZK.id
      maxPrice = [double]$xml.DocumentElement.fcsNotificationZK.lot.maxPrice
      maxPriceStr = ([double]$xml.DocumentElement.fcsNotificationZK.lot.maxPrice).ToString('C')
      purchaseObjectInfo = $xml.DocumentElement.fcsNotificationZK.purchaseObjectInfo
      href = $xml.DocumentElement.fcsNotificationZK.href
      attachments = $attachments
      attachmentsStr = "$($attachments.Length):`r`n$(String-Join "`r`n" ($attachments|%{$_.fileName}))"
      dateOpen = $dateOpen
      dateOpenStr = $dateOpen.ToString("dd.MM.yyyy hh:mm")
      purchaseObjects = $purchaseObjects
      po_OKPD2s = String-Join "`r`n" ($purchaseObjects |%{$_.OKPD2}|select -unique)
      po_OKPD2codes = String-Join "`r`n" ($purchaseObjects |%{$_.OKPD2Code}|select -unique)
      po_names = String-Join "`r`n" ($purchaseObjects |%{$_.name}|select -unique)
      is_Good = $false
      is_Double = $false
   } | DictToObj
   $result.is_Good = filterZK44 $result
   $result
}


function DownloadAttachments{

   ReadZips | ? { $_.is_Good } | % {

      if($glob.DownloadedErrorCounter -gt $glob.MaxDownloadedErrorCounter) { return }

      $path = "$(AttachStore)$($_.regNumber)_$($_.notifyId)"
      if(test-path $path) { return }

      $errorOccured = $false
      $notify = $_


      $attachments = $_.attachments |%{
         try {
            $url = $_.url
            [System.Threading.Thread]::Sleep(1000)
            @{
               content = (Invoke-WebRequest -Method "get" -Uri $url).Content
               fileName = $_.filename
            }
         }
         catch {
            $glob.DownloadedErrorCounter++
            $errorOccured=$true

            write-host
            write-host "$($glob.DownloadedErrorCounter)`: $($notify.zipFile)"
            write-host $notify.reqFile
            write-host "$url $_"
         }
      }|?{$_.Content}

      if($errorOccured) { return }
      mkdir $path |out-null
      $attachments|%{
         [System.IO.File]::WriteAllBytes("$path\$($_.fileName)", $_.Content)
      }
      $_|ConvertTo-Json -Depth 50 |Out-File "$path\info.txt"
      readSrcXml $_ |Out-File "$path\info.xml"
   }
}

function readSrcXml ($notify)
{
   $zip = [System.IO.Compression.ZipFile]::Open("$(FtpStore)$($notify.zipFile)","Read")
   $entry = $zip.Entries |?{$_.FullName -match $notify.reqFile} | select -first 1 
   StreamToString $entry.Open()
}

function ClearAttachments {
   if($glob.DownloadedErrorCounter > 0){return}

   $goods = ReadZips |?{$_.is_Good}|%{"$($_.regNumber)_$($_.notifyId)"}

   dir "$(AttachStore)*" |%{
      if($_.Name.split('_').Length -ne 2) { continue }

      if(!$goods.Contains($_.Name))
      {
         rd $_.FullName -Force -Recurse
      }

   }
}

function analizeOKPD2 { ReadZips|%{$_.purchaseObjects}| select -Property OKPD2Code,OKPD2 |Group-Object OKPD2 |select -Property count,@{Name="OKPD2";Expression={($_.group |select -first 1 ).OKPD2Code}},name}

function makeHtm {
   @'
   <html>
   <head>
      <meta charset="utf-8"/>
      <meta content="width=device-width, initial-scale=1, shrink-to-fit=no" name="viewport"/>
      <link crossorigin="anonymous" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" rel="stylesheet"/>
   </head>
   <body onload='bodyLoad();'>
      <script src="https://cdnjs.cloudflare.com/ajax/libs/handlebars.js/3.0.3/handlebars.js"></script>
      <script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.1/jquery.min.js"></script>
      <!-- <script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>-->
      <script crossorigin="anonymous" integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q" src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js"></script>
      <script crossorigin="anonymous" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js"></script>
      <script>
         var obj = {"reg":
'@
   (ReadZips | ?{$_.is_Good} |sort -property dateOpen | ConvertTo-Json -Depth 50 -compress)
@'
   };

         function bodyLoad()
         {
            var template = Handlebars.compile($("#div-template").html(), {
                 noEscape: true
               });
            $("#insert-pos").html(template(obj));
         }
      </script>

      <div id="insert-pos"></div>

      <script id="div-template" type="text/x-handlebars-template">
         <div class="container-fluid">
            {{#each reg}}
            <div class="row">
               <div class="col-3">
                  <h10><a href="{{href}}">N {{regNumber}}</a></h10>
                  <h5>{{maxPriceStr}}</h5>
                  <p>{{dateOpenStr}}</p>
                  <p>   <a class="btn btn-primary"
                        href="attachments/{{regNumber}}_{{notifyId}}"
                        title="{{#each attachments}}{{fileName}}
 {{/each}}">
                        ATTACH
                     </a>           
                  </p>       
               </div>
               <div class="col-9">
                  <h5>{{purchaseObjectInfo}}</h5>
                  <div class="container-fluid">
                     {{#each purchaseObjects}}
                     <div class="row">
                        <div class="col">{{name}}</div>
                        <div class="col">
                           {{quantity}} {{edism}} x {{priceStr}} ={{sumStr}}
                        </div>
                        <div class="col">
                           <a href="http://help-tender.ru/okpd2.asp?id={{OKPD2Code}}">{{OKPD2Code}}</a>
                           {{OKPD2}}
                        </div>
                     </div>
                     {{/each}}
                  </div>
               </div>
            </div>
            {{/each}}
        </div>
      </script>
   </body>
</html>
'@
}

function _Do {
   "init"
   Init
   "SaveFtpZips"
   SaveFtpZips
   "ClearFtpZips"
   ClearFtpZips
   # ReadZips |select -Property is_Good,regNumber,notifyId,maxPrice,dateOpen,purchaseObjectInfo,po_names,po_OKPD2s,po_OKPD2codes,attachmentsStr,ZipFile,is_Double | ogv
   # analizeOKPD2 |ogv
   "DownloadAttachments"
   DownloadAttachments
   #ClearAttachments
   "makeHtm"
   makeHtm |Out-File "$($glob.dir)index.htm" -Encoding UTF8

}

_Do
<#
TODO
инфа для подачи заявки
окдп2 фильтр все неподходят, а не одна
разделитель между заявками
фильтр на дату создания
ссылки открыть в новой вкладке
#>
