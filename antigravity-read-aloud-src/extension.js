const vscode = require('vscode');
const { spawn, exec } = require('child_process');
const fs = require('fs');
const path = require('path');
const os = require('os');

let currentProcess = null;
let currentDecorationType = null;

function activate(context) {
    // Creamos el estilo de resaltado (ej: fondo amarillo semitransparente)
    currentDecorationType = vscode.window.createTextEditorDecorationType({
        backgroundColor: 'rgba(255, 255, 0, 0.3)',
        borderRadius: '2px'
    });

    let disposableSpeak = vscode.commands.registerCommand('antigravity-read-aloud.speakSelection', async function () {
        let text = "";
        const editor = vscode.window.activeTextEditor;
        let startOffset = 0;

        if (editor) {
            let selection = editor.selection;
            if (!selection.isEmpty) {
                text = editor.document.getText(selection);
                startOffset = editor.document.offsetAt(selection.start);
            } else {
                text = editor.document.getText();
            }
        }

        let isFallback = false;
        if (!text || !text.trim()) {
            try {
                text = await vscode.env.clipboard.readText();
                isFallback = true;
            } catch (err) {
                console.error("Error leyendo portapapeles:", err);
            }

            if (!text || !text.trim()) {
                vscode.window.showInformationMessage('No hay texto para leer. Selecciona texto en un editor o cópialo (Ctrl+C).');
                return;
            }
        }

        stopReading(editor);

        vscode.window.showInformationMessage('Leyendo texto en voz alta (Presiona Ctrl+Shift+S para detener)...');

        const tempDir = os.tmpdir();
        const textFilePath = path.join(tempDir, 'antigravity-tts.txt');
        const scriptFilePath = path.join(tempDir, 'antigravity-tts.ps1');

        fs.writeFileSync(textFilePath, text, 'utf8');

        // Script PowerShell que emite la posición actual
        const psScript = `
Add-Type -AssemblyName System.Speech
$speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
$text = Get-Content -Path "${textFilePath}" -Raw -Encoding UTF8

Register-ObjectEvent -InputObject $speak -EventName SpeakProgress -Action {
    $pos = $Event.SourceEventArgs.CharacterPosition
    $len = $Event.SourceEventArgs.CharacterCount
    [Console]::WriteLine("POS:$pos,$len")
} | Out-Null

$prompt = $speak.SpeakAsync($text)
while (-not $prompt.IsCompleted) {
    Start-Sleep -Milliseconds 50
}
`;
        fs.writeFileSync(scriptFilePath, psScript, 'utf8');

        // IMPORTANTE: Al usar spawn, usar opción shell: true en Windows o comandos explícitos para no trancar el proceso de la extensión
        currentProcess = spawn('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden', '-File', scriptFilePath]);

        if (editor && !isFallback) {
            currentProcess.stdout.on('data', (data) => {
                const textOutput = data.toString();
                const lines = textOutput.split(/\r?\n/);
                for (const line of lines) {
                    const match = line.match(/^POS:(\d+),(\d+)/);
                    if (match) {
                        const wordPos = parseInt(match[1], 10);
                        const wordLen = parseInt(match[2], 10);

                        const absoluteStart = startOffset + wordPos;
                        const absoluteEnd = absoluteStart + wordLen;

                        try {
                            const startPos = editor.document.positionAt(absoluteStart);
                            const endPos = editor.document.positionAt(absoluteEnd);
                            const range = new vscode.Range(startPos, endPos);

                            editor.setDecorations(currentDecorationType, [range]);
                        } catch (e) { }
                    }
                }
            });

            currentProcess.stderr.on('data', (data) => {
                console.error("TTS Error:", data.toString());
            });
        }

        currentProcess.on('close', () => {
            currentProcess = null;
            if (editor) {
                editor.setDecorations(currentDecorationType, []);
            }
        });
    });

    let disposableStop = vscode.commands.registerCommand('antigravity-read-aloud.stopSpeaking', function () {
        stopReading(vscode.window.activeTextEditor);
        vscode.window.showInformationMessage('Lectura detenida.');
    });

    context.subscriptions.push(disposableSpeak);
    context.subscriptions.push(disposableStop);
}

function stopReading(editor) {
    if (editor && currentDecorationType) {
        editor.setDecorations(currentDecorationType, []);
    }
    if (currentProcess) {
        currentProcess.kill();
        currentProcess = null;
    }
}

function deactivate() {
    stopReading(null);
}

module.exports = {
    activate,
    deactivate
}
