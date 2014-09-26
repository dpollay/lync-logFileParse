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
            MainVideoCodec, MainVideoCaptureDevName, MainVideoCaptureDevDriver, MainVideoRenderDevName, MainVideoRenderDevDriver

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
            $information.MainVideoCodec = $MainVideo.InboundStream.Payload.Video.PayloadDescription[1]
            $information.MainVideoCaptureDevName = $MainVideo.Description.CaptureDev.Name[1]
            $information.MainVideoCaptureDevDriver = $MainVideo.Description.CaptureDev.Driver[1]
            $information.MainVideoRenderDevName = $MainVideo.Description.RenderDev.Name[1]
            $information.MainVideoRenderDevDriver = $MainVideo.Description.RenderDev.Driver[1]
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
        @{name="Main Video Codec";expression={$information.MainMainVideoCodec}}, `
        @{name="Main Video Capture Name";expression={$information.MainVideoCaptureDevName}}, `
        @{name="Main Video Capture Driver";expression={$information.MainVideoCaptureDevDriver}}, `
        @{name="Main Video Render Name";expression={$information.MainVideoRenderDevName}}, `
        @{name="Main Video Render Driver";expression={$information.MainVideoRenderDevDriver}} `
        | ConvertTo-Html -head $header | Out-File $htmlfile -append

Invoke-Expression $htmlfile