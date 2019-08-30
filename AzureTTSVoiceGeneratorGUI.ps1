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
V0.01, 23/08/2019 - Initial version
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

#region XAML window definition
# Right-click XAML and choose WPF/Edit... to edit WPF Design
# in your favorite WPF editing tool
$xaml = @'
<Window
   xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
   xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
   Width ="750"
   SizeToContent="WidthAndHeight"
   Title="AzureTTSVoiceGenerator" Height="430" ResizeMode="CanMinimize" ShowInTaskbar="False" WindowStartupLocation="CenterScreen" MinWidth="750" MinHeight="430">
    <Grid Margin="10,10,10,0" Height="387" VerticalAlignment="Top">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBox x:Name="Box_TextMessage" HorizontalAlignment="Left" Height="123" Margin="10,254,-453.5,-206" TextWrapping="Wrap" Text="Place here the text you want to convert to a voice message" VerticalAlignment="Top" Width="595"/>
        <Button x:Name="Button_Run" Content="RUN!" HorizontalAlignment="Left" Margin="616,254,-559.5,-206" VerticalAlignment="Top" Width="95" Height="123"/>
        <ComboBox x:Name="ComboBox_Location" HorizontalAlignment="Left" Margin="98,14,-262.5,0" VerticalAlignment="Top" Width="331" IsEditable="True" IsSynchronizedWithCurrentItem="True">
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
        <TextBox x:Name="Box_Key" HorizontalAlignment="Left" Height="23" Margin="98,45,-263.5,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="331"/>
        <Label Content="Output Path" HorizontalAlignment="Left" Margin="10,91,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.484,0.464"/>
        <Label Content="Output File" HorizontalAlignment="Left" Margin="10,123,0,0" VerticalAlignment="Top" Width="74"/>
        <TextBox x:Name="Box_Output_Path" HorizontalAlignment="Left" Height="23" Margin="98,92,-264.5,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="331"/>
        <TextBox x:Name="Box_Output_File" HorizontalAlignment="Left" Height="23" Margin="98,124,-266.5,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="331"/>
        <Label Content="Audio Format" HorizontalAlignment="Left" Margin="10,187,0,-41" VerticalAlignment="Top" Width="115"/>
        <Label Content="Voice" HorizontalAlignment="Left" Margin="10,217,0,-71" VerticalAlignment="Top" Width="115"/>
        <ComboBox x:Name="ComboBox_Format" HorizontalAlignment="Left" Margin="98,189,-261.5,-32" VerticalAlignment="Top" Width="331" IsSynchronizedWithCurrentItem="True">
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
        <ComboBox x:Name="ComboBox_Voice" HorizontalAlignment="Left" Margin="98,219,-558.5,-69" VerticalAlignment="Top" Width="613" IsSynchronizedWithCurrentItem="True"/>
        <Button x:Name="Button_Save" Content="Save Settings" HorizontalAlignment="Left" Margin="616,14,-564.5,0" VerticalAlignment="Top" Width="95" Height="22"/>
        <Button x:Name="Button_Reload" Content="Reload Settings" HorizontalAlignment="Left" Margin="616,46,-563.5,0" VerticalAlignment="Top" Width="95" Height="22"/>
        <Label Content="Output" HorizontalAlignment="Left" Margin="10,152,0,0" VerticalAlignment="Top" Width="84"/>
        <Label x:Name="Label_Output" Content="" HorizontalAlignment="Left" Margin="98,152,-556.5,0" VerticalAlignment="Top" Width="613"/>
        <Button x:Name="Button_Browse" Content="Browse" HorizontalAlignment="Left" Margin="444,92,-354.5,0" VerticalAlignment="Top" Width="75" Height="23"/>
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
  
 <UserControl x:Class="FolderBrowserDialogServiceSample.Views.FolderBrowserDialogView" 
    ...
    xmlns:dxmvvm="http://schemas.devexpress.com/winfx/2008/xaml/mvvm">
    <dxmvvm:Interaction.Behaviors>
        <dxmvvm:FolderBrowserDialogService />
    </dxmvvm:Interaction.Behaviors>
    ...
</UserControl>
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
