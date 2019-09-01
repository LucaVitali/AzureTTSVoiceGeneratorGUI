<#
Azure TTS Voice Generator
Azure Cognitive Services Text to Speech

.SYNOPSIS
AzureTTSVoiceGeneratorGUI.ps1

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
V1.00, 31/08/2019 - Initial version
#>

#region InitializeVariables
$ScriptPath = $MyInvocation.MyCommand.Path
$ConfigFile = ([System.IO.Path]::ChangeExtension($ScriptPath, "xml")) 
$Key = ""
$Location = ""
$AudioPath = ""
$AudioFile = ""
$AudioFormat = ""
$Voice = ""
#endregion InitializeVariables

Function ReadSettings ()
{
	if (Test-Path -Path "$($ConfigFile)")
	{
		try
		{
			$xml = [xml](get-Content -path "$($ConfigFile)")
			$Key = $xml.configuration.SavedKey
			$Location = $xml.configuration.SavedLocation
			$AudioPath = $xml.configuration.SavedAudioPath
      $AudioFile = $xml.configuration.SavedAudioFile
      $AudioFormat = $xml.configuration.SavedAudioFormat
      $Voice = $xml.configuration.SavedVoice
		}
		catch
		{
      $Key = ""
      $Location = ""
			$AudioPath = ""
      $AudioFile = ""
      $AudioFormat = ""
      $Voice = ""
		}
	}
	else
	{
      $Key = ""
      $Location = ""
			$AudioPath = ""
      $AudioFile = ""
      $AudioFormat = ""
      $Voice = ""
	}
	return $Key,$Location,$AudioPath,$AudioFile,$AudioFormat,$Voice
}

Function WriteSettings ()
{
	param ([string]$myConfigFile, [string]$Key, [string]$Location, [string]$AudioPath, [string]$AudioFile, [string]$AudioFormat, [string]$Voice)
		[xml]$Doc = New-Object System.Xml.XmlDocument
		$Dec = $Doc.CreateXmlDeclaration("1.0","UTF-8",$null)
		$Doc.AppendChild($Dec) | out-null
		$Root = $Doc.CreateNode("element","configuration",$null)
		$Element = $Doc.CreateElement("SavedKey")
		$Element.InnerText = $Key
		$Root.AppendChild($Element) | out-null
		$Element = $Doc.CreateElement("SavedLocation")
		$Element.InnerText = $Location
		$Root.AppendChild($Element) | out-null
		$Element = $Doc.CreateElement("SavedAudioPath")
		$Element.InnerText = $AudioPath
		$Root.AppendChild($Element) | out-null
		$Element = $Doc.CreateElement("SavedAudioFile")
		$Element.InnerText = $AudioFile
		$Root.AppendChild($Element) | out-null
		$Element = $Doc.CreateElement("SavedAudioFormat")
		$Element.InnerText = $AudioFormat
		$Root.AppendChild($Element) | out-null
		$Element = $Doc.CreateElement("SavedVoice")
		$Element.InnerText = $Voice
		$Root.AppendChild($Element) | out-null
		$Doc.AppendChild($Root) | out-null
		try
		{
			$Doc.save(("$($myConfigFile)"))
		}
		catch
		{
		}
}

function Get-Folder {
[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[string] $Message = "Select the desired folder",
		[int] $path = 0x00
	)
  [Object] $FolderObject = New-Object -ComObject Shell.Application
  $folder = $FolderObject.BrowseForFolder(0, $message, 0, $path)
  if ($folder -ne $null) {
		return $folder.self.Path
  } else {
  	Write-Host "No folder specified"
  }
}

#region XAML window definition
# Right-click XAML and choose WPF/Edit... to edit WPF Design
# in your favorite WPF editing tool
$xaml = @'
<Window
   xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
   xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
   Width ="750"
   SizeToContent="WidthAndHeight"
   Title="AzureTTSVoiceGeneratorGUI" Height="550" ResizeMode="CanMinimize" ShowInTaskbar="False" WindowStartupLocation="CenterScreen" MinWidth="750" MinHeight="550">
    <Grid Margin="10,10,10,0" Height="508" VerticalAlignment="Top">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBox x:Name="Box_TextMessage" HorizontalAlignment="Left" Height="123" Margin="10,319,-436.5,-262" TextWrapping="Wrap" Text="Place here the text you want to convert to a voice message" VerticalAlignment="Top" Width="595"/>
        <Button x:Name="Button_Run" Content="RUN!" HorizontalAlignment="Left" Margin="616,319,-542.5,-262" VerticalAlignment="Top" Width="95" Height="123"/>
        <ComboBox x:Name="ComboBox_Location" HorizontalAlignment="Left" Margin="98,14,-275,0" VerticalAlignment="Top" Width="421" IsEditable="True" IsSynchronizedWithCurrentItem="True">
            <ComboBoxItem Content="australiaeast"/>
            <ComboBoxItem Content="canadacentral"/>
            <ComboBoxItem Content="centralus"/>
            <ComboBoxItem Content="eastasia"/>
            <ComboBoxItem Content="eastus"/>
            <ComboBoxItem Content="eastus2"/>
            <ComboBoxItem Content="francecentral"/>
            <ComboBoxItem Content="centralindia"/>
            <ComboBoxItem Content="japaneast"/>
            <ComboBoxItem Content="koreacentral"/>
            <ComboBoxItem Content="northcentralus"/>
            <ComboBoxItem Content="northeurope"/>
            <ComboBoxItem Content="southcentralus"/>
            <ComboBoxItem Content="southeastasia"/>
            <ComboBoxItem Content="uksouth"/>
            <ComboBoxItem Content="westeurope"/>
            <ComboBoxItem Content="westus"/>
            <ComboBoxItem Content="westus2"/>
        </ComboBox>
        <Label Content="Location" HorizontalAlignment="Left" Margin="10,12,0,0" VerticalAlignment="Top"/>
        <Label Content="Key" HorizontalAlignment="Left" Margin="10,44,0,0" VerticalAlignment="Top" Width="49"/>
        <TextBox x:Name="Box_Key" HorizontalAlignment="Left" Height="22" Margin="98,46,-276,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="421"/>
        <Label Content="Output Folder" HorizontalAlignment="Left" Margin="10,147,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.484,0.464"/>
        <Label Content="Output File" HorizontalAlignment="Left" Margin="10,179,0,-24" VerticalAlignment="Top" Width="74"/>
        <TextBox x:Name="Box_Output_Path" HorizontalAlignment="Left" Height="23" Margin="98,148,-263,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="421"/>
        <TextBox x:Name="Box_Output_File" HorizontalAlignment="Left" Height="23" Margin="98,180,-264,-21" TextWrapping="Wrap" VerticalAlignment="Top" Width="421"/>
        <Label Content="Audio Format" HorizontalAlignment="Left" Margin="10,252,0,-98" VerticalAlignment="Top" Width="115"/>
        <Label Content="Voice" HorizontalAlignment="Left" Margin="10,282,0,-128" VerticalAlignment="Top" Width="115"/>
        <ComboBox x:Name="ComboBox_Format" HorizontalAlignment="Left" Margin="98,254,-262,-93" VerticalAlignment="Top" Width="421" IsSynchronizedWithCurrentItem="True">
            <ComboBoxItem Content="raw-16khz-16bit-mono-pcm"/>
            <ComboBoxItem Content="raw-8khz-8bit-mono-mulaw"/>
            <ComboBoxItem Content="riff-8khz-8bit-mono-alaw"/>
            <ComboBoxItem Content="riff-8khz-8bit-mono-mulaw"/>
            <ComboBoxItem Content="riff-16khz-16bit-mono-pcm" FontWeight="Bold"/>
            <ComboBoxItem Content="audio-16khz-128kbitrate-mono-mp3"/>
            <ComboBoxItem Content="audio-16khz-64kbitrate-mono-mp3"/>
            <ComboBoxItem Content="audio-16khz-32kbitrate-mono-mp3"/>
            <ComboBoxItem Content="raw-24khz-16bit-mono-pcm"/>
            <ComboBoxItem Content="riff-24khz-16bit-mono-pcm"/>
            <ComboBoxItem Content="audio-24khz-160kbitrate-mono-mp3"/>
            <ComboBoxItem Content="audio-24khz-96kbitrate-mono-mp3"/>
            <ComboBoxItem Content="audio-24khz-48kbitrate-mono-mp3"/>
        </ComboBox>
        <ComboBox x:Name="ComboBox_Voice" HorizontalAlignment="Left" Margin="98,284,-214,-95" VerticalAlignment="Top" Width="421" IsSynchronizedWithCurrentItem="True"/>
        <Button x:Name="Button_Save" Content="Save Settings" HorizontalAlignment="Left" Margin="616,14,-564.5,0" VerticalAlignment="Top" Width="95" Height="22"/>
        <Button x:Name="Button_Reload" Content="Reload Settings" HorizontalAlignment="Left" Margin="616,46,-540.5,0" VerticalAlignment="Top" Width="95" Height="22"/>
        <Label Content="Output" HorizontalAlignment="Left" Margin="10,208,0,-53" VerticalAlignment="Top" Width="84"/>
        <Label x:Name="Label_Output" Content="" HorizontalAlignment="Left" Margin="98,208,-460,-53" VerticalAlignment="Top" Width="613"/>
        <Button x:Name="Button_Browse" Content="Browse" HorizontalAlignment="Left" Margin="616,148,-455,0" VerticalAlignment="Top" Width="95" Height="23"/>
        <Label Content="Token Service Endpoint" HorizontalAlignment="Left" Margin="10,80,0,0" VerticalAlignment="Top" Width="133"/>
        <Label Content="Cognitive Services TTS Endpoint" HorizontalAlignment="Left" Margin="10,111,0,0" VerticalAlignment="Top" Width="183"/>
        <Label x:Name="Label_Token_URI" Content="" HorizontalAlignment="Left" Margin="193,80,-464,0" VerticalAlignment="Top" Width="521"/>
        <Label x:Name="Label_Service_URI" Content="" HorizontalAlignment="Left" Margin="193,111,-461,0" VerticalAlignment="Top" Width="518"/>
        <Label Content="Be careful: existing files with &#xD;&#xA;the same name will be &#xD;&#xA;overwritten without any alert!" HorizontalAlignment="Left" Margin="531,179,-407,-35" Width="180" Height="66" VerticalAlignment="Top"/>
        <Label Content="If available, prefer neural &#xD;&#xA;voices to standard ones" HorizontalAlignment="Left" Margin="531,260,-362,-59" VerticalAlignment="Top" Height="42" Width="180"/>
        <TextBlock HorizontalAlignment="Left" Margin="474,465,-354,-119" TextWrapping="Wrap" VerticalAlignment="Top" Height="41" Width="237"><Run Text="[LinkedIn]"/><Run Text=" "/><Hyperlink NavigateUri="https://linkedin.com/in/lucavitali"><Run Text="https://linkedin.com/in/lucavitali"/></Hyperlink><LineBreak/><Run Text="[Github]"/><Run Text="&#x9;  "/><Hyperlink NavigateUri="https://github.com/LucaVitali"><Run Text="https://github.com/LucaVitali"/></Hyperlink></TextBlock>
        <TextBlock HorizontalAlignment="Left" Margin="10,452,-28,-64" TextWrapping="Wrap" VerticalAlignment="Top" Width="376"><Run Text="Created by Luca Vitali - Microsoft Office Apps &amp; Services MVP"/><LineBreak/><Run Text="[Blog] &#x9;  "/><Hyperlink NavigateUri="https://lucavitali.wordpress.com/"><Run Text="https://lucavitali.wordpress.com"/></Hyperlink><LineBreak/><Run Text="[Twitter] &#x9;  "/><Hyperlink NavigateUri="https://twitter.com/Luca_Vitali"><Run Text="https://twitter.com/Luca_Vitali"/></Hyperlink></TextBlock>
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
Add-Type -AssemblyName presentationCore
$Location = $window.ComboBox_Location.Text
$AudioPath = $window.Box_Output_Path.Text
$AudioFile = $window.Box_Output_File.Text
$AudioFormat = $window.ComboBox_Format.Text
$RequestHeaders = @{"Authorization"= $OAuthToken;"Content-Type"= "application/ssml+xml";"X-Microsoft-OutputFormat"= $AudioFormat;"User-Agent" = "MIMText2Speech"}
[xml]$VoiceBody = @"
<speak version='1.0' xmlns="http://www.w3.org/2001/10/synthesis" xml:lang='en-US'> 
  <voice  name='$($window.ComboBox_Voice.Text)'>
    VoiceMessage
  </voice>
</speak>
"@
 $VoiceBody.speak.voice.'#text' = $window.Box_TextMessage.Text
 Invoke-RestMethod -Method POST -Uri $ServiceURI -Headers $RequestHeaders -Body $VoiceBody -ContentType "application/ssml+xml" -OutFile "$($AudioPath)\$($AudioFile)"
}

$window.Button_Save.add_Click{
  # remove param() block if access to event information is not required
  param
  (
    [Parameter(Mandatory)][Object]$sender,
    [Parameter(Mandatory)][Windows.RoutedEventArgs]$e
  )
  
  WriteSettings $Configfile $window.Box_Key.Text $window.ComboBox_Location.Text $window.Box_Output_Path.Text $window.Box_Output_File.Text $window.ComboBox_Format.Text $window.ComboBox_Voice.Text
}

$window.Button_Reload.add_Click{
  # remove param() block if access to event information is not required
  param
  (
    [Parameter(Mandatory)][Object]$sender,
    [Parameter(Mandatory)][Windows.RoutedEventArgs]$e
  )
  $Key,$Location,$AudioPath,$AudioFile,$AudioFormat,$Voice = ReadSettings
  $window.Box_Key.Text= $Key
  $window.ComboBox_Location.Text= $Location
  $window.Box_Output_Path.Text= $AudioPath
  $window.Box_Output_File.Text= $AudioFile
  $window.ComboBox_Format.Text= $AudioFormat
  $TokenURI = "https://$($location).api.cognitive.microsoft.com/sts/v1.0/issueToken"
  $ServiceURI = "https://$($Location).tts.speech.microsoft.com/cognitiveservices/v1"
  $TokenHeaders = @{"Content-type"= "application/x-www-form-urlencoded";"Content-Length"= "0";"Ocp-Apim-Subscription-Key"= $Key}
  $OAuthToken = Invoke-RestMethod -Method POST -Uri $TokenURI -Headers $TokenHeaders
  $VoiceListURI = "https://$($Location).tts.speech.microsoft.com/cognitiveservices/voices/list"
  $RequestHeadersGET = @{"Authorization"= $OAuthToken}
  $VoiceList = Invoke-RestMethod -Method GET -Uri $VoiceListURI -Headers $RequestHeadersGET
  $window.ComboBox_Voice.Text= $Voice
 }

$window.Box_Output_Path.add_SelectionChanged{
  # remove param() block if access to event information is not required
  param
  (
    [Parameter(Mandatory)][Object]$sender,
    [Parameter(Mandatory)][Windows.RoutedEventArgs]$e
  )
  
  $window.Label_Output.Content= "$($window.Box_Output_Path.Text)\$($window.Box_Output_File.Text)"
}

$window.Box_Output_File.add_SelectionChanged{
  # remove param() block if access to event information is not required
  param
  (
    [Parameter(Mandatory)][Object]$sender,
    [Parameter(Mandatory)][Windows.RoutedEventArgs]$e
  )
  
  $window.Label_Output.Content= "$($window.Box_Output_Path.Text)\$($window.Box_Output_File.Text)"
}

$window.Button_Browse.add_Click{
  # remove param() block if access to event information is not required
  param
  (
    [Parameter(Mandatory)][Object]$sender,
    [Parameter(Mandatory)][Windows.RoutedEventArgs]$e
  )
  $window.Box_Output_Path.Text = (Get-Folder "Select the output folder or create a new one")
}

# Show Window
$Key,$Location,$AudioPath,$AudioFile,$AudioFormat,$Voice = ReadSettings
$TokenURI = "https://$($location).api.cognitive.microsoft.com/sts/v1.0/issueToken"
$ServiceURI = "https://$($Location).tts.speech.microsoft.com/cognitiveservices/v1"
$TokenHeaders = @{"Content-type"= "application/x-www-form-urlencoded";"Content-Length"= "0";"Ocp-Apim-Subscription-Key"= $Key}
$OAuthToken = Invoke-RestMethod -Method POST -Uri $TokenURI -Headers $TokenHeaders
$VoiceListURI = "https://$($Location).tts.speech.microsoft.com/cognitiveservices/voices/list"
$RequestHeadersGET = @{"Authorization"= $OAuthToken}
$VoiceList = Invoke-RestMethod -Method GET -Uri $VoiceListURI -Headers $RequestHeadersGET
$window.Box_Key.Text= $Key
$window.ComboBox_Location.Text= $Location
$window.Box_Output_Path.Text= $AudioPath
$window.Box_Output_File.Text= $AudioFile
$window.ComboBox_Format.Text= $AudioFormat
$window.ComboBox_Voice.ItemsSource= $VoiceList.Name
$window.ComboBox_Voice.Text= $Voice
$window.Label_Output.Content= "$($window.Box_Output_Path.Text)\$($window.Box_Output_File.Text)"
$result = Show-WPFWindow -Window $window
#region Process results
if ($result -eq $true)
{

}
else
{

}
#endregion Process results
