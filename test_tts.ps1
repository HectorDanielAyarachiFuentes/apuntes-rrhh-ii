
Add-Type -AssemblyName System.Speech
$speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
$text = "Prueba de audio desde el editor sincrono"

Register-ObjectEvent -InputObject $speak -EventName SpeakProgress -Action {
    $pos = $Event.SourceEventArgs.CharacterPosition
    $len = $Event.SourceEventArgs.CharacterCount
    Write-Host "POS:$pos,$len"
} | Out-Null

$speak.Speak($text)
