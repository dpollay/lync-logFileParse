<#
Lync Log File Parser
Copyright (C) 2014 David Pollay
This program is free software; you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation; either version 2 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License along
with this program; if not, write to the Free Software Foundation, Inc.,
51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.

#>
function Lync-ParseLog{

Begin{
    $htmlfile = $env:HOME + '\LyncLog.html'
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
<center>
<title>
Lync Log File Parsing
</title>
"@
}

Process{ $endpointstats = select-String $env:LocalAppData\Microsoft\Office\15.0\Lync\Tracing\*.UccApilog* -pattern "<vqreport" |
    ForEach-Object {
        $check =  $_.Line -match "</VQReportEvent>"
        if($check -NotMatch "True"){
            $_.Line += "</VQSessionReport></VQReportEvent>"
        }
        $xmldoc = [xml]$_.Line
        $information = $_ | Select-Object -Property Name, OS, CPUName, CPUNumberOfCores
        $endpointChoices = "Name", "OS", "CPUName", "CPUNumberOfCores"
        $endpoint = $xmldoc.VQReportEvent.VQSessionReport.Endpoint
        $information.Name = $endpoint.Name
        $information.OS = $endpoint.OS
        $information.CPUName = $endpoint.CPUName
        $information.CPUNumberOfCores = $endpoint.CPUNumberOfCores
        $information
    } | select @{name="Computer Name";expression={$information.Name}}, `
    @{name="Operating System";expression={$information.OS}}, `
    @{name="CPU Name";expression={$information.CPUName}}, `
    @{name="# of CPU Cores";expression={$information.CPUNumberOfCores}} `
    | ConvertTo-Html -Fragment -As Table -PreContent "<h2>Computer Information</h2>" | Out-String

$dialoginfostats = select-String $env:LocalAppData\Microsoft\Office\15.0\Lync\Tracing\*.UccApilog* -pattern "<vqreport" |
    ForEach-Object {
        $check =  $_.Line -match "</VQReportEvent>"
        if($check -NotMatch "True"){
            $_.Line += "</VQSessionReport></VQReportEvent>"
        }
        $xmldoc = [xml]$_.Line
        $information = $_ | Select-Object -Property Start, End, FromURI, ToURI, LocalUserAgent
        $dialog = $xmldoc.VQReportEvent.VQSessionReport.DialogInfo
        $information.Start = $dialog.Start
        $information.End = $dialog.End
        $information.FromURI = $dialog.FromURI
        $information.ToURI = $dialog.ToURI
        $information.LocalUserAgent = $dialog.LocalUserAgent
        $information
    } | select @{name="Start Time";expression={[datetime]$information.Start}}, `
    @{name="End Time";expression={[datetime]$information.End}}, `
    @{name="From User";expression={$information.FromURI}}, `
    @{name="To User";expression={$information.ToURI}}, `
    @{name="Lync Version";expression={$information.LocalUserAgent}} `
    | ConvertTo-Html -Fragment -As Table -PreContent "<h2>Call Setup Information</h2>" | Out-String

$audiostats = select-String $env:LocalAppData\Microsoft\Office\15.0\Lync\Tracing\*.UccApilog* -pattern "<vqreport" |
    ForEach-Object {
        $check =  $_.Line -match "</VQReportEvent>"
        if($check -NotMatch "True"){
            $_.Line += "</VQSessionReport></VQReportEvent>"
        }
        $xmldoc = [xml]$_.Line
        $information = $_ | Select-Object -Property AudioCodec, AudioCaptureDevName, AudioCaptureDevDriver, AudioRenderDevName, AudioRenderDevDriver, RecvListenMOS
        if($xmldoc.VQReportEvent.VQSessionReport.DialogInfo.ToURI -notmatch "applicationsharing"){
            $mainaudio = $xmldoc.VQReportEvent.VQSessionReport.MediaLine
            $information.AudioCodec = $mainaudio.InboundStream.Payload.Audio.PayloadDescription
            $information.AudioCaptureDevName = $mainaudio.Description.CaptureDev.Name[0]
            $information.AudioCaptureDevDriver = $mainaudio.Description.CaptureDev.Driver[0]
            $information.AudioRenderDevName = $mainaudio.Description.RenderDev.Name[0]
            $information.AudioRenderDevDriver = $mainaudio.Description.RenderDev.Driver[0]
            $information.RecvListenMOS = $mainaudio.InboundStream.QualityEstimates.Audio.RecvListenMOS
        }
        else{
            $information.AudioCodec = "N/A"
            $information.AudioCaptureDevName = "N/A"
            $information.AudioCaptureDevDriver = "N/A"
            $information.AudioRenderDevName = "N/A"
            $information.AudioRenderDevDriver = "N/A"
            $information.RecvListenMOS = "N/A"
        }
        $information
    } | select @{name="Codec";expression={$information.AudioCodec}}, `
    @{name="Capture Device";expression={$information.AudioCaptureDevName}}, `
    @{name="Capture Driver";expression={$information.AudioCaptureDevDriver}}, `
    @{name="Render Device";expression={$information.AudioRenderDevName}}, `
    @{name="Render Driver";expression={$information.AudioRenderDevDriver}}, `
    @{name="MOS";expression={$information.RecvListenMOS}} `
    | ConvertTo-Html -Fragment -As Table -PreContent "<h2>Audio Information</h2>" | Out-String

$videostats = select-String $env:LocalAppData\Microsoft\Office\15.0\Lync\Tracing\*.UccApilog* -pattern "<vqreport" |
    ForEach-Object {
        $check =  $_.Line -match "</VQReportEvent>"
        if($check -NotMatch "True"){
            $_.Line += "</VQSessionReport></VQReportEvent>"
        }
        $xmldoc = [xml]$_.Line
        $information = $_ | Select-Object -Property MainVideoCodec, MainVideoCaptureDevName, MainVideoCaptureDevDriver, MainVideoRenderDevName, MainVideoRenderDevDriver, mainVideoResolution, mainVideoFrameRateAvg, `
            mainVideoCifQuality, mainVideoVGAQuality, mainVideoHDQuality
        if($xmldoc.VQReportEvent.VQSessionReport.DialogInfo.ToURI -notmatch "applicationsharing"){
            $MainVideo = $xmldoc.VQReportEvent.VQSessionReport.MediaLine
            $information.MainVideoCodec = $MainVideo.InboundStream.Payload.Video.RecvCodecTypes[0]
            $information.MainVideoCaptureDevName = $MainVideo.Description.CaptureDev.Name[1]
            $information.MainVideoCaptureDevDriver = $MainVideo.Description.CaptureDev.Driver[1]
            $information.MainVideoRenderDevName = $MainVideo.Description.RenderDev.Name[1]
            $information.MainVideoRenderDevDriver = $MainVideo.Description.RenderDev.Driver[1]
            $information.MainVideoResolution = $MainVideo.InboundStream.Payload.Video.Resolution[0]
            $information.MainVideoFrameRateAvg = $MainVideo.InboundStream.Payload.Video.VideoFrameRateAvg[0]
            $information.MainVideoCifQuality = $MainVideo.InboundStream.Payload.Video.VideoResolutionDistribution.CIFQuality[0]
            $information.MainVideoVGAQuality = $MainVideo.InboundStream.Payload.Video.VideoResolutionDistribution.VGAQuality[0]
            $information.MainVideoHDQuality = $MainVideo.InboundStream.Payload.Video.VideoResolutionDistribution.HD720Quality[0]
        }
        else{
            $information.MainVideoCodec = "N/A"
            $information.MainVideoCaptureDevName = "N/A"
            $information.MainVideoCaptureDevDriver = "N/A"
            $information.MainVideoRenderDevName = "N/A"
            $information.MainVideoRenderDevDriver = "N/A"
            $information.MainVideoResolution = "N/A"
            $information.MainVideoFrameRateAvg = "N/A"
            $information.MainVideoCifQuality = "N/A"
            $information.MainVideoVGAQuality = "N/A"
            $information.MainVideoHDQuality = "N/A"
        }
        $information
    } | select @{name="Codec";expression={$information.MainVideoCodec}}, `
    @{name="Capture Device";expression={$information.MainVideoCaptureDevName}}, `
    @{name="Capture Driver";expression={$information.MainVideoCaptureDevDriver}}, `
    @{name="Render Device";expression={$information.MainVideoRenderDevName}}, `
    @{name="Render Driver";expression={$information.MainVideoRenderDevDriver}}, `
    @{name="Avg FPS";expression={$information.MainVideoFrameRateAvg}}, `
    @{name="Resolution";expression={$information.MainVideoResolution}}, `
    @{name="% CIF Quality";expression={$information.MainVideoCifQuality}}, `
    @{name="% VGA Quality";expression={$information.MainVideoVGAQuality}}, `
    @{name="% HD Quality";expression={$information.MainVideoHDQuality}} `
    | ConvertTo-Html -Fragment -As Table -PreContent "<h2>Video Information</h2>" | Out-String

$Report = ConvertTo-Html -Title "Lync Log File Parsing" `
    -Head $header `
    -Body "$endpointstats $dialoginfostats $audiostats $videostats"
}

End{$report | out-file $htmlfile ; Invoke-Expression $htmlfile}

}
Lync-ParseLog