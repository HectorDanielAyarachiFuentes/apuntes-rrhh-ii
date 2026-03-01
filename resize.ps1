Add-Type -AssemblyName System.Drawing
$sourcePath = 'C:\Users\Ramoncito\.gemini\antigravity\brain\e99a4c35-1a3e-4535-a5d2-2648f4cacee4\premium_tab_icon_1772287463270.png'
$destDir = 'c:\Users\Ramoncito\.antigravity\tablero-menu-ayarachi-2.0\Menu-tab-2.0.0\icons'

$sizes = @(128, 48, 16)
$img = [System.Drawing.Image]::FromFile($sourcePath)

foreach ($s in $sizes) {
    $destPath = Join-Path $destDir "icon$s.png"
    $bmp = New-Object System.Drawing.Bitmap($s, $s)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $rect = New-Object System.Drawing.Rectangle(0, 0, $s, $s)
    $g.DrawImage($img, $rect)
    $bmp.Save($destPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose()
    $bmp.Dispose()
}
$img.Dispose()
