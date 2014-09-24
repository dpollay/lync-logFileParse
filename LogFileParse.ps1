$htmlfile = '.\LyncLog.html'
out-file $htmlfile

$Header = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
.odd  { background-color:#ffffff; }
.even { background-color:#dddddd; }
TR:Nth-Child(Even) {Background-Color: #dddddd;}
TR:Hover TD {Background-Color: #C1D5F8;}
</style>
<title>
Lync Log File Parse
</title>
"@

select-String $env:LocalAppData\Microsoft\Office\15.0\Lync\Tracing\*.UccApilog* -pattern "<vqreport" |
    ForEach-Object {
        $information = $_ | Select-Object -Property Name, v2:CPUName, v2:CPUNumberOfCores, LocalUserAgent, Start, End, ToURI, RecvListenMOS, LossRate, PayloadDescription, v2:CIFQuality, v2:VGAQuality, v2:HD720Quality, VideoFrameRateAvg, Resolution, v3:RecvCodecTypes

        $headerProperties = 'Name', 'v2:CPUName', 'v2:CPUNumberOfCores', 'Start', 'End', 'LocalUserAgent'

        foreach($headerproperty in $headerProperties){
            $infoHeaderTemp = ($_.Line -split '" ')
            $information.$headerproperty = ($infoHeaderTemp -split '">') | select-string $headerproperty= | foreach-object {
                $_ -replace '="', ' ' ` -replace $headerproperty, '' ` -replace ' ','' 
            }
        }

        $InboundAudioProperties = 'ToURI', 'RecvListenMOS', 'LossRate', 'PayloadDescription'

        foreach($inboundAudioProperty in $inboundAudioProperties){
            $infoInboundAudioTempPre = ($_.Line -split '/InboundStream>')
            $infoInboundAudioTemp = ($infoInboundAUdioTempPre[0] -split '/>')
            $information.$inboundAudioProperty = ($infoInboundAudioTemp -split '><') |select-string $inboundAudioProperty'>'|foreach-object {
                $_ -replace '>', ' ' ` -replace $inboundAudioProperty, '' ` -replace '</', '' 
            }
         }

         $InboundVideoProperties = 'v2:CIFQuality', 'v2:VGAQuality', 'v2:HD720Quality', 'VideoFrameRateAvg', 'Resolution', 'v3:RecvCodecTypes'

         foreach($inboundVideoProperty in $InboundVideoProperties){
            $infoInboundVideoTemp = $infoInboundAudioTempPre[1] ` -split '<InboundStream' ` -split '/InboundStream>'
            $information.$inboundVideoProperty = ($infoInboundVideoTemp[1] -split '><') | select-string $inboundVideoProperty'>' | foreach-object {
                $_ -replace '>', ' ' ` -replace $inboundVideoProperty, '' -replace '</', '' 
            }
         }
    $information 
 #   $pre += "<b>Computer Name:  " + $information.Name[0] + `
 #       "<br>CPU:  " + $information."v2:CPUName" + `
 #       "<br>CPU Cores:  " + $information."v2:CPUNumberOfCores" + `
 #       "<br>Lync Version:  " + $information.LocalUserAgent + `
 #       "<br><br></b>"
 #   $post += "End File"
    } | select @{name="Start Time";expression={$_.Start}}, `
        @{name="End Time";expression={$_.End}}, `
        @{name="To User";expression={$_.ToURI}}, `
        @{name="MOS";expression={$_.RecvListenMOS}}, `
        @{name="Packet Loss";expression={$_.LossRate}}, `
        @{name="Audio Codec";expression={$_.PayloadDescription}}, `
        @{name="CIF Quality %";expression={$_."v2:CIFQuality"}}, `
        @{name="VGA Quality %";expression={$_."v2:VGAQuality"}}, `
        @{name="HD Quality %";expression={$_."v2:HD720Quality"}}, `
        @{name="FPS";expression={$_.VideoFrameRateAvg}}, `
        @{name="Resolution";expression={$_.Resolution}}, `
        @{name="Video Codec";expression={$_."v3:RecvCodecTypes"}} `
        | ConvertTo-Html -head $header| Out-File $htmlfile -append

Invoke-Expression $htmlfile