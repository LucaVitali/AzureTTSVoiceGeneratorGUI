<#
Azure TTS Voice Generator GUI
Azure Cognitive Services Text to Speech

.SYNOPSIS
AzureTTSVoiceGeneratorGUI.ps1

.DESCRIPTION 
PowerShell script to generate Voice Messages with Azure Cognitive Services Text to Speech
Quick Link: http://bit.ly/AzureTTSGUI

.NOTES
Written by: Luca Vitali - Microsoft Office Apps & Services MVP

Find me on:
[Blog]		https://lucavitali.wordpress.com/
[Twitter]	https://twitter.com/Luca_Vitali
[LinkedIn]	https://www.linkedin.com/in/lucavitali/
[GitHub]	https://github.com/LucaVitali

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
V1.00, 02/09/2019 - Initial version
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
$xaml = @'
<Window
   xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
   xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
   Width ="670"
   SizeToContent="WidthAndHeight"
   Title="AzureTTSVoiceGeneratorGUI" Height="655" ResizeMode="CanMinimize" ShowInTaskbar="False" WindowStartupLocation="CenterScreen" MinWidth="670" MinHeight="655">
    <Grid Margin="10,10,4,0" Height="612" VerticalAlignment="Top">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Rectangle HorizontalAlignment="Left" Height="83" Margin="10,307,0,0" Stroke="Gray" VerticalAlignment="Top" Width="625"/>
        <Rectangle HorizontalAlignment="Left" Height="118" Margin="10,165,-118,0" Stroke="Gray" VerticalAlignment="Top" Width="625"/>
        <Rectangle HorizontalAlignment="Left" Height="127" Margin="10,14,-123,0" Stroke="Gray" VerticalAlignment="Top" Width="625"/>
        <TextBox x:Name="Box_TextMessage" HorizontalAlignment="Left" Height="111" Margin="10,401,0,0" TextWrapping="Wrap" Text="Place here the text you want to convert to a voice message" VerticalAlignment="Top" Width="509"/>
        <Button x:Name="Button_Run" Content="RUN!" HorizontalAlignment="Left" Margin="531,401,0,0" VerticalAlignment="Top" Width="95" Height="44"/>
        <ComboBox x:Name="ComboBox_Location" HorizontalAlignment="Left" Margin="98,35,-2,0" VerticalAlignment="Top" Width="411" IsEditable="True" IsSynchronizedWithCurrentItem="True">
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
        <Label Content="Location" HorizontalAlignment="Left" Margin="11,31,0,0" VerticalAlignment="Top"/>
        <Label Content="Key" HorizontalAlignment="Left" Margin="11,58,0,0" VerticalAlignment="Top" Width="49"/>
        <TextBox x:Name="Box_Key" HorizontalAlignment="Left" Height="22" Margin="98,62,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="411"/>
        <Label Content="Output Folder" HorizontalAlignment="Left" Margin="10,176,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.484,0.464"/>
        <Label Content="Output File" HorizontalAlignment="Left" Margin="10,208,0,0" VerticalAlignment="Top" Width="74"/>
        <TextBox x:Name="Box_Output_Path" HorizontalAlignment="Left" Height="22" Margin="98,178,-6,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="421"/>
        <TextBox x:Name="Box_Output_File" HorizontalAlignment="Left" Height="22" Margin="98,210,-6,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="421"/>
        <Label Content="Audio Format" HorizontalAlignment="Left" Margin="11,317,0,0" VerticalAlignment="Top" Width="83"/>
        <Label Content="Voice" HorizontalAlignment="Left" Margin="10,354,0,0" VerticalAlignment="Top" Width="83"/>
        <ComboBox x:Name="ComboBox_Format" HorizontalAlignment="Left" Margin="98,321,0,0" VerticalAlignment="Top" Width="421" IsSynchronizedWithCurrentItem="True">
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
        <ComboBox x:Name="ComboBox_Voice" HorizontalAlignment="Left" Margin="98,358,0,0" VerticalAlignment="Top" Width="421" IsSynchronizedWithCurrentItem="True"/>
        <Button x:Name="Button_Save" Content="Save Settings" HorizontalAlignment="Left" Margin="531,458,0,0" VerticalAlignment="Top" Width="95" Height="22"/>
        <Button x:Name="Button_Reload" Content="Reload Settings" HorizontalAlignment="Left" Margin="531,490,0,0" VerticalAlignment="Top" Width="95" Height="22"/>
        <Label Content="Output" HorizontalAlignment="Left" Margin="10,238,0,0" VerticalAlignment="Top" Width="84"/>
        <Label x:Name="Label_Output" Content="" HorizontalAlignment="Left" Margin="98,238,-113,0" VerticalAlignment="Top" Width="528"/>
        <Button x:Name="Button_Browse" Content="Browse" HorizontalAlignment="Left" Margin="531,178,-113,0" VerticalAlignment="Top" Width="95" Height="22"/>
        <Label Content="Token Service" HorizontalAlignment="Left" Margin="11,89,0,0" VerticalAlignment="Top" Width="84"/>
        <Label Content="TTS Endpoint" HorizontalAlignment="Left" Margin="11,115,0,0" VerticalAlignment="Top" Width="84"/>
        <Label x:Name="Label_Token_URI" Content="" HorizontalAlignment="Left" Margin="98,89,0,0" VerticalAlignment="Top" Width="411"/>
        <Label x:Name="Label_Service_URI" Content="" HorizontalAlignment="Left" Margin="98,115,0,0" VerticalAlignment="Top" Width="411"/>
        <Label Content="Be careful: existing file with the same Output name will be overwritten without any alert!" HorizontalAlignment="Left" Margin="10,258,0,0" Width="521" Height="25" VerticalAlignment="Top"/>
        <Label Content="Azure Cognitive Services " HorizontalAlignment="Left" Margin="22,0,0,0" VerticalAlignment="Top" Width="142" Background="White"/>
        <TextBox HorizontalAlignment="Left" Height="33" Margin="524,317,0,0" TextWrapping="Wrap" Text="Suggested format in Bold" VerticalAlignment="Top" Width="104" BorderBrush="{x:Null}"/>
        <Label Content="Destination" HorizontalAlignment="Left" Margin="22,152,0,0" VerticalAlignment="Top" Width="71" Background="White"/>
        <Label Content="Audio" HorizontalAlignment="Left" Margin="22,293,0,0" VerticalAlignment="Top" Width="43" Background="White"/>
        <TextBlock HorizontalAlignment="Left" Margin="10,517,0,0" TextWrapping="Wrap" VerticalAlignment="Top" FontSize="16"><Run Text="Azure Text to Speech Voice Generator GUI"/><LineBreak/><Run Text="Created by Luca Vitali - Microsoft Office Apps &amp; Services MVP"/></TextBlock>
        <Button x:Name="Link_Blog" Content="https://lucavitali.wordpress.com" HorizontalAlignment="Left" Margin="63,561,0,0" VerticalAlignment="Top" Width="181" Height="20" BorderBrush="{x:Null}" Background="{x:Null}" HorizontalContentAlignment="Left" Padding="0" VerticalContentAlignment="Top" Foreground="#FF0066CC" />
        <Label Content="[Blog]" HorizontalAlignment="Left" Height="20" Margin="10,561,0,0" VerticalAlignment="Top" Width="48" Padding="0" RenderTransformOrigin="0.507,0.125"/>
        <Label Content="[Twitter]" HorizontalAlignment="Left" Height="20" Margin="10,581,0,0" VerticalAlignment="Top" Width="48" Padding="0"/>
        <Button x:Name="Link_Twitter" Content="https://twitter.com/Luca_Vitali" HorizontalAlignment="Left" Margin="63,581,0,0" VerticalAlignment="Top" Width="181" Height="20" BorderBrush="{x:Null}" Background="{x:Null}" HorizontalContentAlignment="Left" Padding="0" VerticalContentAlignment="Top" Foreground="#FF0066CC" />
        <Label Content="[LinkedIn]" HorizontalAlignment="Left" Margin="278,561,0,43" Width="55" Padding="0" RenderTransformOrigin="0.507,0.125"/>
        <Label Content="[GitHub]" HorizontalAlignment="Left" Height="20" Margin="278,581,0,0" VerticalAlignment="Top" Width="48" Padding="0"/>
        <Button x:Name="Link_GitHub" Content="https://github.com/LucaVitali" HorizontalAlignment="Left" Margin="338,581,0,0" VerticalAlignment="Top" Width="181" Height="20" BorderBrush="{x:Null}" Background="{x:Null}" HorizontalContentAlignment="Left" Padding="0" VerticalContentAlignment="Top" Foreground="#FF0066CC" />
        <Button x:Name="Link_LinkedIn" Content="https://linkedin.com/in/lucavitali" HorizontalAlignment="Left" Margin="338,561,0,0" VerticalAlignment="Top" Width="181" Height="20" BorderBrush="{x:Null}" Background="{x:Null}" HorizontalContentAlignment="Left" Padding="0" VerticalContentAlignment="Top" Foreground="#FF0066CC" />
        <Button x:Name="Link_TTS_Services" Content="Learn more about&#xD;&#xA;Azure TTS Services" HorizontalAlignment="Left" Margin="524,101,-75,0" VerticalAlignment="Top" Width="102" Height="35" BorderBrush="{x:Null}" Background="{x:Null}" HorizontalContentAlignment="Left" Padding="0" VerticalContentAlignment="Center" Foreground="#FF0066CC" />
        <Button x:Name="Link_Create_Account" Content="How to create a &#xD;&#xA;free Azure TTS &#xD;&#xA;Account" HorizontalAlignment="Left" Margin="524,35,-70,0" VerticalAlignment="Top" Width="102" Height="49" BorderBrush="{x:Null}" Background="{x:Null}" HorizontalContentAlignment="Left" Padding="0" VerticalContentAlignment="Center" Foreground="#FF0066CC" />
        <TextBlock HorizontalAlignment="Left" Margin="531,530,0,0" TextWrapping="Wrap" Text="Version 1.0" VerticalAlignment="Top" Width="95" FontSize="16" TextAlignment="Center"/>
        <Button x:Name="Link_Check_Update" Content=" Check&#xD;&#xA;Update" HorizontalAlignment="Left" Margin="531,551,0,0" VerticalAlignment="Top" Width="95" Height="44" BorderBrush="{x:Null}" Background="{x:Null}" HorizontalContentAlignment="Center" Padding="0" VerticalContentAlignment="Center" Foreground="#FF0066CC" FontSize="16" />
        <TextBox HorizontalAlignment="Left" Height="33" Margin="524,354,0,0" TextWrapping="Wrap" Text="Prefer neural voices&#xD;&#xA;to standard ones" VerticalAlignment="Top" Width="111" BorderBrush="{x:Null}"/>
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

$window.ComboBox_Location.add_LostFocus{
  # remove param() block if access to event information is not required
  param
  (
    [Parameter(Mandatory)][Object]$sender,
    [Parameter(Mandatory)][Windows.RoutedEventArgs]$e
  )
   $Key= $window.Box_Key.Text
   $Location= $window.ComboBox_Location.Text
   $window.Label_Token_URI.Content= "https://$($window.ComboBox_Location.Text).api.cognitive.microsoft.com/sts/v1.0/issueToken"
   $window.Label_Service_URI.Content= "https://$($window.ComboBox_Location.Text).tts.speech.microsoft.com/cognitiveservices/v1"
   $TokenURI = "https://$($location).api.cognitive.microsoft.com/sts/v1.0/issueToken"
   $ServiceURI = "https://$($Location).tts.speech.microsoft.com/cognitiveservices/v1"
   $TokenHeaders = @{"Content-type"= "application/x-www-form-urlencoded";"Content-Length"= "0";"Ocp-Apim-Subscription-Key"= $Key}
   $OAuthToken = Invoke-RestMethod -Method POST -Uri $TokenURI -Headers $TokenHeaders
   $VoiceListURI = "https://$($Location).tts.speech.microsoft.com/cognitiveservices/voices/list"
   $RequestHeadersGET = @{"Authorization"= $OAuthToken}
   $VoiceList = Invoke-RestMethod -Method GET -Uri $VoiceListURI -Headers $RequestHeadersGET
   $window.ComboBox_Voice.Text= $Voice
}

$window.Link_Blog.add_Click{
  # remove param() block if access to event information is not required
  param
  (
    [Parameter(Mandatory)][Object]$sender,
    [Parameter(Mandatory)][Windows.RoutedEventArgs]$e
  )
  Start-Process ("https://lucavitali.wordpress.com");
}

$window.Link_Twitter.add_Click{
  # remove param() block if access to event information is not required
  param
  (
    [Parameter(Mandatory)][Object]$sender,
    [Parameter(Mandatory)][Windows.RoutedEventArgs]$e
  )
  Start-Process ("https://twitter.com/Luca_Vitali");
}

$window.Link_LinkedIn.add_Click{
  # remove param() block if access to event information is not required
  param
  (
    [Parameter(Mandatory)][Object]$sender,
    [Parameter(Mandatory)][Windows.RoutedEventArgs]$e
  )
  Start-Process ("https://linkedin.com/in/lucavitali");
}

$window.Link_GitHub.add_Click{
  # remove param() block if access to event information is not required
  param
  (
    [Parameter(Mandatory)][Object]$sender,
    [Parameter(Mandatory)][Windows.RoutedEventArgs]$e
  )
  Start-Process ("https://github.com/LucaVitali");
}

$window.Link_TTS_Services.add_Click{
  # remove param() block if access to event information is not required
  param
  (
    [Parameter(Mandatory)][Object]$sender,
    [Parameter(Mandatory)][Windows.RoutedEventArgs]$e
  )
  Start-Process ("https://docs.microsoft.com/en-us/azure/cognitive-services/speech-service/rest-text-to-speech");
}

$window.Link_Create_Account.add_Click{
  # remove param() block if access to event information is not required
  param
  (
    [Parameter(Mandatory)][Object]$sender,
    [Parameter(Mandatory)][Windows.RoutedEventArgs]$e
  )
  Start-Process ("http://bit.ly/AzureTTSGUI");
}

$window.Link_Check_Update.add_Click{
  # remove param() block if access to event information is not required
  param
  (
    [Parameter(Mandatory)][Object]$sender,
    [Parameter(Mandatory)][Windows.RoutedEventArgs]$e
  )
  Start-Process ("https://github.com/LucaVitali/AzureTTSVoiceGeneratorGUI");
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
$window.Label_Token_URI.Content= "https://$($window.ComboBox_Location.Text).api.cognitive.microsoft.com/sts/v1.0/issueToken"
$window.Label_Service_URI.Content= "https://$($window.ComboBox_Location.Text).tts.speech.microsoft.com/cognitiveservices/v1"
$result = Show-WPFWindow -Window $window
#region Process results
if ($result -eq $true)
{

}
else
{

}
#endregion Process results
# SIG # Begin signature block
# MIINBAYJKoZIhvcNAQcCoIIM9TCCDPECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUjisDTJ4klGxB04r1pS+I/iPX
# yfOgggpGMIIFDjCCA/agAwIBAgIQBOtO2q9gdCIMsGwltOx52jANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTE5MDkyMzAwMDAwMFoXDTIwMDky
# MzEyMDAwMFowSzELMAkGA1UEBhMCSVQxEDAOBgNVBAcTB0JvbG9nbmExFDASBgNV
# BAoTC0x1Y2EgVml0YWxpMRQwEgYDVQQDEwtMdWNhIFZpdGFsaTCCASIwDQYJKoZI
# hvcNAQEBBQADggEPADCCAQoCggEBANXCDXqLmdYBnUvNYwgwW812mz/tS4kzZBof
# dN8bmz4wpDmZu8wOB5FAtKwGKh0iRJqWGS+Q1KAWmxQAfk3pa18dGUuwXErkVSWv
# da/uGI4xp/Pa8dGJEW0731iwVv0+mIfVCvNnvwNMtCUHVREdhqIY9hPt+tMwRISr
# 3Y5vdcbGwk8a6XChNJ25zWvkeM/PXrhh+HfyZ0S5H1UpoL5MY3aFjUDX7/v40Ik7
# GRTv45J7KzNAzr76pmdZ62D6WqMiG/tzpPSW3IsEEVcCw+OC4Rh+yC9FtBJYc33X
# 2m7XdGjaOHNRgc170dPRGTSLK8iwUDE5Pao4yDjH+inznMbURlECAwEAAaOCAcUw
# ggHBMB8GA1UdIwQYMBaAFFrEuXsqCqOl6nEDwGD5LfZldQ5YMB0GA1UdDgQWBBRk
# uwaRC4jr8ftOUu8KTo9OACw4FjAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYI
# KwYBBQUHAwMwdwYDVR0fBHAwbjA1oDOgMYYvaHR0cDovL2NybDMuZGlnaWNlcnQu
# Y29tL3NoYTItYXNzdXJlZC1jcy1nMS5jcmwwNaAzoDGGL2h0dHA6Ly9jcmw0LmRp
# Z2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3MtZzEuY3JsMEwGA1UdIARFMEMwNwYJ
# YIZIAYb9bAMBMCowKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNv
# bS9DUFMwCAYGZ4EMAQQBMIGEBggrBgEFBQcBAQR4MHYwJAYIKwYBBQUHMAGGGGh0
# dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBOBggrBgEFBQcwAoZCaHR0cDovL2NhY2Vy
# dHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0U0hBMkFzc3VyZWRJRENvZGVTaWduaW5n
# Q0EuY3J0MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggEBAJmPbVC/Na5L
# B76NcBoHd9hJblom+SLPe7cRzDk5zy96aP0BD3fzq176SOIirqdBl/sS8fxVRZ8/
# ys9SmdSAYD+yJK1liBKyVvTV/Jx7jF8QLvu8L8C/7ebcUzMjeIT3jNpgm3hCeiDT
# rNX71JM+xXTc7E3rVhXNG2alfi36MD/I1GD21cS5I4FBlascomWH6ui8+5GsGQOo
# 5RL6NDr1YFe13dbdii921lQTgeVGxi3DO4WifCJ+7fzJbKeUDrJGISBIAX8aDaU9
# 2mYpWQV81EtbB8s3bCWo6Yn6CUhbVkfvD6YIHqWVeJlnwzqImJI8b0wasJ9WQOS+
# q/2BESeaGzIwggUwMIIEGKADAgECAhAECRgbX9W7ZnVTQ7VvlVAIMA0GCSqGSIb3
# DQEBCwUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAX
# BgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNVBAMTG0RpZ2lDZXJ0IEFzc3Vy
# ZWQgSUQgUm9vdCBDQTAeFw0xMzEwMjIxMjAwMDBaFw0yODEwMjIxMjAwMDBaMHIx
# CzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3
# dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJ
# RCBDb2RlIFNpZ25pbmcgQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQD407Mcfw4Rr2d3B9MLMUkZz9D7RZmxOttE9X/lqJ3bMtdx6nadBS63j/qSQ8Cl
# +YnUNxnXtqrwnIal2CWsDnkoOn7p0WfTxvspJ8fTeyOU5JEjlpB3gvmhhCNmElQz
# UHSxKCa7JGnCwlLyFGeKiUXULaGj6YgsIJWuHEqHCN8M9eJNYBi+qsSyrnAxZjNx
# PqxwoqvOf+l8y5Kh5TsxHM/q8grkV7tKtel05iv+bMt+dDk2DZDv5LVOpKnqagqr
# hPOsZ061xPeM0SAlI+sIZD5SlsHyDxL0xY4PwaLoLFH3c7y9hbFig3NBggfkOItq
# cyDQD2RzPJ6fpjOp/RnfJZPRAgMBAAGjggHNMIIByTASBgNVHRMBAf8ECDAGAQH/
# AgEAMA4GA1UdDwEB/wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDAzB5BggrBgEF
# BQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBD
# BggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0
# QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2Ny
# bDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDig
# NoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9v
# dENBLmNybDBPBgNVHSAESDBGMDgGCmCGSAGG/WwAAgQwKjAoBggrBgEFBQcCARYc
# aHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAKBghghkgBhv1sAzAdBgNVHQ4E
# FgQUWsS5eyoKo6XqcQPAYPkt9mV1DlgwHwYDVR0jBBgwFoAUReuir/SSy4IxLVGL
# p6chnfNtyA8wDQYJKoZIhvcNAQELBQADggEBAD7sDVoks/Mi0RXILHwlKXaoHV0c
# LToaxO8wYdd+C2D9wz0PxK+L/e8q3yBVN7Dh9tGSdQ9RtG6ljlriXiSBThCk7j9x
# jmMOE0ut119EefM2FAaK95xGTlz/kLEbBw6RFfu6r7VRwo0kriTGxycqoSkoGjpx
# KAI8LpGjwCUR4pwUR6F6aGivm6dcIFzZcbEMj7uo+MUSaJ/PQMtARKUT8OZkDCUI
# QjKyNookAv4vcn4c10lFluhZHen6dGRrsutmQ9qzsIzV6Q3d9gEgzpkxYz0IGhiz
# gZtPxpMQBvwHgfqL2vmCSfdibqFT+hKUGIUukpHqaGxEMrJmoecYpJpkUe8xggIo
# MIICJAIBATCBhjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5j
# MRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBT
# SEEyIEFzc3VyZWQgSUQgQ29kZSBTaWduaW5nIENBAhAE607ar2B0IgywbCW07Hna
# MAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3
# DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEV
# MCMGCSqGSIb3DQEJBDEWBBT1t/xh/XBFvh00xr9oC5v0+xC/YjANBgkqhkiG9w0B
# AQEFAASCAQAODfzCqPSQoIEcDkN8+xVNvA533PFj5JtCNVtD7I8ZUGIpjEKkvWXz
# 6LRUogHCBkeW+3ICIO1LaACQLvsvlHHEaDOa7Fkf2gJVsvMs6TASMlp30b0foyjq
# e+TPTvCpRer+BwbWLVNfRzHrzKoLCA4XqTeyigP8zcBeUddCA9qOrZLub34DW95m
# q9mCI8Q4HPhS/nfSPsfKU/Y8D6slr0QC8Z31UODpyXO1st2pqsfJv5lbGVtQFV+b
# AsuGZUj6ZZGCHUMzDYcusgdQc8eg6n008ChRYCnA65DHvRc/i1fM2Z+iZyaPBFrW
# 8Iib2bKp6vzjymr6LS8f8PQFPrL6NbSY
# SIG # End signature block
