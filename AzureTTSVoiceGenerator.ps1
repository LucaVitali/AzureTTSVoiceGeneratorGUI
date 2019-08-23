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

# Note: Trial keys will be for the 'westus' location only



# Generate Request Auth Header
$Location = "westus"
$TokenURI = "https://$($location).api.cognitive.microsoft.com/sts/v1.0/issueToken"
$Key1 = "YOUR_SUBSCRIPTION_KEY"
$TokenHeader = @{"Ocp-Apim-Subscription-Key" = $Key1;"Content-Length"= "0";"Content-type" = "application/x-www-form-urlencoded"}
            
# Get OAuth Token
$OAuthToken = Invoke-RestMethod -Method POST -Uri $TokenURI -Headers $TokenHeader


# Output Settings
Add-Type -AssemblyName presentationCore

# Output Path
$AudioPath = "C:\temp\"

# Output File
$AudioFile = "audiooutexample.mp3"


# Text to Speech Endpoint
$URI = "https://$($location).tts.speech.microsoft.com/cognitiveservices/v1"

# Output formats
#ssml-16khz-16bit-mono-tts 
#raw-16khz-16bit-mono-pcm 
#audio-16khz-16kbps-mono-siren 
#riff-16khz-16kbps-mono-siren 
#riff-16khz-16bit-mono-pcm 
#audio-16khz-128kbitrate-mono-mp3 
#audio-16khz-64kbitrate-mono-mp3 
#audio-16khz-32kbitrate-mono-mp3

$Headers = @{"Ocp-Apim-Subscription-Key" = $key1; 
            "Content-Type" = "application/ssml+xml";
            "X-Microsoft-OutputFormat" = "audio-16khz-32kbitrate-mono-mp3";
            "User-Agent" = "MIMText2Speech";
            "Authorization" = $OAuthToken
            }

# Voices https://docs.microsoft.com/en-us/azure/cognitive-services/speech/api-reference-rest/bingvoiceoutput#SupLocales 
#Microsoft Server Speech Text to Speech Voice (en-US, JessaRUS)
#Microsoft Server Speech Text to Speech Voice (en-GB, Susan, Apollo)
#Microsoft Server Speech Text to Speech Voice (en-AU, HayleyRUS) 

[xml]$Voice = @'
<speak version='1.0' xmlns="http://www.w3.org/2001/10/synthesis" xml:lang='en-US'> 
  <voice  name='Microsoft Server Speech Text to Speech Voice (en-AU, HayleyRUS)'>
    TEXTTOCONVERT
  </voice>
</speak>
'@

# Inject text to convert
$Voice.speak.voice.'#text' = "I just converted this string to speech using Azure"
$Voice.speak.voice.'#text'  

# Voice Message Generation
Invoke-RestMethod -Method POST -Uri $URI -Headers $Headers -Body $Voice -ContentType "application/ssml+xml" -OutFile "$($AudioPath)$($AudioFile)" 
