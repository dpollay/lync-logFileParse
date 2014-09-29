#################################################################################
#  Lync Log File Parser
#  Copyright (C) 2014 David Pollay
#  This program is free software; you can redistribute it and/or modify
#  it under the terms of the GNU General Public License as published by
#  the Free Software Foundation; either version 2 of the License, or
#  (at your option) any later version.
#
#  This program is distributed in the hope that it will be useful,
#  but WITHOUT ANY WARRANTY; without even the implied warranty of
#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#  GNU General Public License for more details.
#
#  You should have received a copy of the GNU General Public License along
#  with this program; if not, write to the Free Software Foundation, Inc.,
#  51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
#################################################################################

$outfileDir = $env:USERPROFILE
$htmlfile = $env:USERPROFILE + '\LyncLog.html'

out-file $htmlfile

# HTML Formatting
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

# Load Lync log files and grab each line that contains a VQ Report
select-String $env:LocalAppData\Microsoft\Office\15.0\Lync\Tracing\*.UccApilog* -pattern "<vqreport" |
    # Convert output into XML Data
    ForEach-Object {
        $check =  $_.Line -match "</VQReportEvent>"
        if($check -NotMatch "True"){
            $_.Line += "</VQSessionReport></VQReportEvent>"
        }
        $xmldoc = [xml]$_.Line

        # Choose Fields to Grab
        $information = $_ | Select-Object -Property Name, OS, CPUName, CPUNumberOfCores, `
            Start, End, FromURI, ToURI, LocalUserAgent, `
            AUdioCodec, AudioCaptureDevName, AudioCaptureDevDriver, AudioRenderDevName, AudioRenderDevDriver, RecvListenMOS, `
            MainVideoCodec, MainVideoCaptureDevName, MainVideoCaptureDevDriver, MainVideoRenderDevName, MainVideoRenderDevDriver, mainVideoResolution, mainVideoFrameRateAvg, `
            mainVideoCifQuality, mainVideoVGAQuality, mainVideoHDQuality

        # Grab Endpoint Statistics
        $endpointChoices = "Name", "OS", "CPUName", "CPUNumberOfCores"
        $endpoint = $xmldoc.VQReportEvent.VQSessionReport.Endpoint
        foreach($endpointchoice in $endpointChoices){
            $information.$endpointchoice = $endpoint.$endpointchoice
        }

        # Grab DialogInfo Statistics
        $dialogChoices = "Start", "End", "FromURI", "ToURI", "LocalUserAgent"
        $dialog = $xmldoc.VQReportEvent.VQSessionReport.DialogInfo
        foreach($dialogchoice in $dialogChoices){
            $information.$dialogchoice = $dialog.$dialogchoice
        }

        # Grab MediaLine main-audio Statistics
        if($information.ToURI -notmatch "applicationsharing"){
            $mainaudio = $xmldoc.VQReportEvent.VQSessionReport.MediaLine
            $information.AudioCodec = $mainaudio.InboundStream.Payload.Audio.PayloadDescription
            $information.AudioCaptureDevName = $mainaudio.Description.CaptureDev.Name[0]
            $information.AudioCaptureDevDriver = $mainaudio.Description.CaptureDev.Driver[0]
            $information.AudioRenderDevName = $mainaudio.Description.RenderDev.Name[0]
            $information.AudioRenderDevDriver = $mainaudio.Description.RenderDev.Driver[0]
            $information.RecvListenMOS = $mainaudio.InboundStream.QualityEstimates.Audio.RecvListenMOS
        }

        # Grab MediaLine main-MainVideo Statistics
        if($information.ToURI -notmatch "applicationsharing"){
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


        $information 
    } | select @{name="Start Time";expression={[datetime]$information.Start}}, ` 
        @{name="End Time";expression={[datetime]$information.End}}, `
        @{name="From User";expression={$information.FromURI}}, `
        @{name="To User";expression={$information.ToURI}}, `
        @{name="Lync Version";expression={$information.LocalUserAgent}}, `
        @{name="Audio Codec";expression={$information.AudioCodec}}, `
        @{name="Audio Capture Name";expression={$information.AudioCaptureDevName}}, `
        @{name="Audio Capture Driver";expression={$information.AudioCaptureDevDriver}}, `
        @{name="Audio Render Name";expression={$information.AudioRenderDevName}}, `
        @{name="Audio Render Driver";expression={$information.AudioRenderDevDriver}}, `
        @{name="Audio MOS";expression={$information.RecvListenMos}}, `
        @{name="Main Video Codec";expression={$information.MainVideoCodec}}, `
        @{name="Main Video Capture Name";expression={$information.MainVideoCaptureDevName}}, `
        @{name="Main Video Capture Driver";expression={$information.MainVideoCaptureDevDriver}}, `
        @{name="Main Video Render Name";expression={$information.MainVideoRenderDevName}}, `
        @{name="Main Video Render Driver";expression={$information.MainVideoRenderDevDriver}}, `
        @{name="Main Video Resolution";expression={$information.MainVideoResolution}}, `
        @{name="Main Video FPS Avg";expression={$information.MainVideoFrameRateAvg}}, `
        @{name="Main Video % CIF Quality";expression={$information.MainVideoCifQuality}}, `
        @{name="Main Video % VGA Quality";expression={$information.MainVideoVGAQUality}}, `
        @{name="Main Video % HD Quality";expression={$information.MainVideoHDQuality}} `
        | ConvertTo-Html -head $header | Out-File $htmlfile -append

Invoke-Expression $htmlfile