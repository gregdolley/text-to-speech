using namespace System.Speech.Synthesis

Add-Type -AssemblyName System.Speech

function Get-DefaultVoice () {
  $speechSynthesizer = [SpeechSynthesizer]::new()
  $speechSynthesizer.GetInstalledVoices() | Select-Object -First 1
}

function Get-VoiceName ($voice) {
  $voice.VoiceInfo.Name
}

function Speak ($Text, [switch]$Print, [Parameter()]$UseVoiceName) {
  $speechSynthesizer = [SpeechSynthesizer]::new()

  if($Print) { Write-Output $Text }
  if(![string]::IsNullOrEmpty($UseVoiceName)) { $speechSynthesizer.SelectVoice($UseVoiceName) }
  
  $speechSynthesizer.Speak($Text)
}

function Get-InstalledVoices() {
  $speechSynthesizer = [SpeechSynthesizer]::new()
  $speechSynthesizer.GetInstalledVoices()
}

$defaultVoice = Get-DefaultVoice
$defaultVoiceName = Get-VoiceName $defaultVoice
$IrinaVoiceName = ""

Write-Output "The voice named '$($defaultVoiceName)' says: "

Speak -Text "Hello, this is a test PowerShell script for converting text into speech." -Print
Speak -Text "The selection of voices you have installed are: " -Print

$counter = 1
$installedVoices = Get-InstalledVoices
foreach ($voice in $installedVoices)
{
  $info = $voice.VoiceInfo
  $name = $info.Name

  Speak -Text " ${counter}: $name" -Print -UseVoiceName $voice.VoiceInfo.Name#$name
 
  if(Select-String -InputObject $name -Pattern "Irina" -SimpleMatch -Quiet)
  {
    $IrinaVoiceName = $name
  }

  $counter++
}

if($IrinaVoiceName -ne "")
{    
    Speak -Text "Now let's see if the Irina voice can speak Russian." -Print
    Speak -Text "Это проверка того, насколько хорошо говорит русский голос Ирины. Когда она говорит по-английски, ее очень сложно понять." -Print -UseVoiceName $IrinaVoiceName
    # Translation: "This is a test of how good Irina's voice speaks Russian. When she speaks in English, she is very difficult to understand."
    # (At least for me... Her English speaking voice often puts accents on the wrong syllables.)
    Speak -Text "Well, at least that's much better than her English pronunciation!!" -Print
}