using namespace System.Speech.Synthesis

Add-Type -AssemblyName System.Speech

$text = "Hello, this is a test PowerShell script for converting text into speech."

$speechSynthesizer = [SpeechSynthesizer]::new()
$installedVoices = $speechSynthesizer.GetInstalledVoices()
$defaultVoice = $installedVoices | Select-Object -First 1
$defaultVoiceName = $defaultVoice.VoiceInfo.Name
$IrinaVoiceName = ""

Write-Output "The voice named '$($defaultVoiceName)' says: "
Write-Output $text

$speechSynthesizer.SelectVoice($defaultVoiceName)
$speechSynthesizer.Speak($text)

$text = "The selection of voices you have installed are: "
Write-Output $text
$speechSynthesizer.Speak($text)
$counter = 1

foreach ($voice in $installedVoices)
{
  $info = $voice.VoiceInfo
  $name = $info.Name

  Write-Output " ${counter}: $name";

  $speechSynthesizer.SelectVoice($name)
  $speechSynthesizer.Speak($name)
 
  if(Select-String -InputObject $name -Pattern "Irina" -SimpleMatch -Quiet)
  {
    $IrinaVoiceName = $name
  }

  $counter++
}

if($IrinaVoiceName -ne "")
{    
    $text = "Now let's see if the Irina voice can speak Russian."   
    Write-Output $text

    $speechSynthesizer.SelectVoice($defaultVoiceName)
    $speechSynthesizer.Speak($text)

    $text = "Это проверка того, насколько хорошо говорит русский голос Ирины. Когда она говорит по-английски, ее очень сложно понять."
    # Translation: "This is a test of how good Irina's voice speaks Russian. When she speaks in English, she is very difficult to understand."
    # (At least for me... Her English speaking voice often puts accents on the wrong syllables.)
    Write-Output $text
    
    $speechSynthesizer.SelectVoice($IrinaVoiceName) 
    $speechSynthesizer.Speak($text)

    $speechSynthesizer.SelectVoice($defaultVoiceName)
    $speechSynthesizer.Speak("Well, that's much better!!")
}