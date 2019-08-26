<#
Azure TTS Voice Generator
Azure Cognitive Services Text to Speech

.SYNOPSIS
AzureTTSVoiceGenerator.ps1

.DESCRIPTION 
PowerShell script to generate Voice Messages with Azure Cognitive Services Text to Speech

.NOTES
Written by: Luca Vitali

Find me on:
* My Blog:	https://lucavitali.wordpress.com/
* Twitter:	https://twitter.com/Luca_Vitali
* LinkedIn:	https://www.linkedin.com/in/lucavitali/

License:

The MIT License (MIT)

Copyright (c) 2019 Luca Vitali

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

Change Log:
V0.01, 23/08/2019 - Initial version
#>

#################
# Authentication
#################

# Generate Request Auth Header
$Location = "westeurope"
$TokenURI = "https://$($location).api.cognitive.microsoft.com/sts/v1.0/issueToken"
$Key1 = "4357584a1225404eb6bc2d110f60585f"
$TokenHeaders = @{
 "Content-type"= "application/x-www-form-urlencoded";
 "Content-Length"= "0";
 "Ocp-Apim-Subscription-Key"= $Key1
 }
            
# Get OAuth Token
$OAuthToken = Invoke-RestMethod -Method POST -Uri $TokenURI -Headers $TokenHeaders

# Text to Speech Endpoint
$URI = "https://$($location).tts.speech.microsoft.com/cognitiveservices/v1"


#################
# Output
#################

# Output Settings
Add-Type -AssemblyName presentationCore

# Default Output Path
$AudioPath = "C:\Temp\"

# Default Output File
$AudioFile = "VoiceMessage1.wav"

# Output formats
#ssml-16khz-16bit-mono-tts 
#raw-16khz-16bit-mono-pcm 
#audio-16khz-16kbps-mono-siren 
#riff-16khz-16kbps-mono-siren 
#riff-16khz-16bit-mono-pcm 
#audio-16khz-128kbitrate-mono-mp3 
#audio-16khz-64kbitrate-mono-mp3 
#audio-16khz-32kbitrate-mono-mp3
$AudioFormat = "riff-16khz-16bit-mono-pcm"

$RequestHeaders = @{
 "Authorization"= $OAuthToken;
 "Content-Type"= "application/ssml+xml";
 "X-Microsoft-OutputFormat"= $AudioFormat;
 "User-Agent" = "MIMText2Speech" 
 }

[xml]$Voice = @'
<speak version='1.0' xmlns="http://www.w3.org/2001/10/synthesis" xml:lang='it-IT'> 
  <voice  name='Microsoft Server Speech Text to Speech Voice (it-IT, ElsaNeural)'>
    TEXTTOCONVERT
  </voice>
</speak>
'@

#################
# XAML
#################

#region XAML window definition
# Right-click XAML and choose WPF/Edit... to edit WPF Design
# in your favorite WPF editing tool
$xaml = @'
<Window
   xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
   xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
   Width ="750"
   SizeToContent="WidthAndHeight"
   Title="AzureTTSVoiceGenerator" Height="418" ResizeMode="CanMinimize" ShowInTaskbar="False" WindowStartupLocation="CenterScreen" MinWidth="750" MinHeight="418">
    <Grid Margin="10,10,10,0" Height="376" VerticalAlignment="Top">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBox x:Name="Box_TextMessage" HorizontalAlignment="Left" Height="123" Margin="10,243,-497.5,-203" TextWrapping="Wrap" Text="Place here the text you want to convert to a voice message" VerticalAlignment="Top" Width="610"/>
        <Button x:Name="Button_Run" Content="RUN!" HorizontalAlignment="Left" Margin="636,243,-588.5,-203" VerticalAlignment="Top" Width="75" Height="123"/>
        <ComboBox x:Name="ComboBox_Location" HorizontalAlignment="Left" Margin="98,14,-196.5,0" VerticalAlignment="Top" Width="208"/>
        <Label Content="Location" HorizontalAlignment="Left" Margin="10,12,0,0" VerticalAlignment="Top"/>
        <Label Content="Key" HorizontalAlignment="Left" Margin="10,44,0,0" VerticalAlignment="Top" Width="49"/>
        <TextBox x:Name="Box_Key" HorizontalAlignment="Left" Height="23" Margin="98,45,-197.5,0" TextWrapping="Wrap" Text="Enter your Key" VerticalAlignment="Top" Width="208"/>
        <Label Content="Output Path" HorizontalAlignment="Left" Margin="10,98,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.484,0.464"/>
        <Label Content="Output File" HorizontalAlignment="Left" Margin="10,130,-17.5,-4" VerticalAlignment="Top" Width="115"/>
        <TextBox x:Name="Box_Output_Path" HorizontalAlignment="Left" Height="23" Margin="98,99,-199.5,0" TextWrapping="Wrap" Text="C:\Temp" VerticalAlignment="Top" Width="208"/>
        <TextBox x:Name="Box_Output_File" HorizontalAlignment="Left" Height="23" Margin="98,131,-198.5,-2" TextWrapping="Wrap" Text="AudioMessage.wav" VerticalAlignment="Top" Width="208"/>
        <Label Content="Audio Format" HorizontalAlignment="Left" Margin="10,167,-6.5,-33" VerticalAlignment="Top" Width="115"/>
        <Label Content="Voice" HorizontalAlignment="Left" Margin="10,197,-3.5,-61" VerticalAlignment="Top" Width="115"/>
        <ComboBox x:Name="ComboBox_Format" HorizontalAlignment="Left" Margin="98,169,-187.5,-31" VerticalAlignment="Top" Width="208"/>
        <ComboBox x:Name="ComboBox_Voice" HorizontalAlignment="Left" Margin="98,199,-184.5,-59" VerticalAlignment="Top" Width="208"/>
    </Grid>
</Window>
'@
#endregion

#region Code Behind
function Convert-XAMLtoWindow
{
  param
  (
    [Parameter(Mandatory=$true)]
    [string]
    $XAML
  )
  
  Add-Type -AssemblyName PresentationFramework
  
  $reader = [XML.XMLReader]::Create([IO.StringReader]$XAML)
  $result = [Windows.Markup.XAMLReader]::Load($reader)
  $reader.Close()
  $reader = [XML.XMLReader]::Create([IO.StringReader]$XAML)
  while ($reader.Read())
  {
      $name=$reader.GetAttribute('Name')
      if (!$name) { $name=$reader.GetAttribute('x:Name') }
      if($name)
      {$result | Add-Member NoteProperty -Name $name -Value $result.FindName($name) -Force}
  }
  $reader.Close()
  $result
}

function Show-WPFWindow
{
  param
  (
    [Parameter(Mandatory=$true)]
    [Windows.Window]
    $Window
  )
  
  $result = $null
  $null = $window.Dispatcher.InvokeAsync{
    $result = $window.ShowDialog()
    Set-Variable -Name result -Value $result -Scope 1
  }.Wait()
  $result
}
#endregion Code Behind

#region Convert XAML to Window
$window = Convert-XAMLtoWindow -XAML $xaml 
#endregion

#region Define Event Handlers
# Right-Click XAML Text and choose WPF/Attach Events to
# add more handlers

$window.Button_Run.add_Click{
  # remove param() block if access to event information is not required
  param
  (
    [Parameter(Mandatory)][Object]$sender,
    [Parameter(Mandatory)][Windows.RoutedEventArgs]$e
  )
 
 
 $AudioPath = $window.Box_Output_Path.Text
 $AudioFile = $window.Box_Output_File.Text
 $Voice.speak.voice.'#text' = $window.Box_TextMessage.Text
 Invoke-RestMethod -Method POST -Uri $URI -Headers $RequestHeaders -Body $Voice -ContentType "application/ssml+xml" -OutFile "$($AudioPath)\$($AudioFile)" 
}

#endregion Event Handlers

# Show Window
$result = Show-WPFWindow -Window $window

#region Process results
if ($result -eq $true)
{
  [PSCustomObject]@{
    EmployeeName = $window.TxtName.Text
    EmployeeMail = $window.TxtEmail.Text
  }
}
else
{
  Write-Warning 'User aborted dialog.'
}
#endregion Process results
