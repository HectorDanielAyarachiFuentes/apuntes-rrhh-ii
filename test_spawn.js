const { spawn } = require('child_process');
const fs = require('fs');

const psScript = `
Add-Type -AssemblyName System.Speech
$speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
$text = "Prueba de audio desde el editor sincrono"

Register-ObjectEvent -InputObject $speak -EventName SpeakProgress -Action {
    $pos = $Event.SourceEventArgs.CharacterPosition
    $len = $Event.SourceEventArgs.CharacterCount
    Write-Host "POS:$pos,$len"
} | Out-Null

$speak.Speak($text)
`;

fs.writeFileSync('test_tts.ps1', psScript);

const p = spawn('powershell.exe', ['-ExecutionPolicy', 'Bypass', '-File', 'test_tts.ps1']);

p.stdout.on('data', d => console.log('OUT:', d.toString()));
p.stderr.on('data', d => console.error('ERR:', d.toString()));
p.on('close', code => console.log('EXIT:', code));
