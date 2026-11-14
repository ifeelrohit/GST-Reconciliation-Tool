Add-Type -AssemblyName System.Drawing
Get-ChildItem -Path "screenshots" -Filter "*.bmp" | ForEach-Object {
    $bmpPath = $_.FullName
    $pngPath = [System.IO.Path]::ChangeExtension($bmpPath, ".png")
    $bmp = [System.Drawing.Bitmap]::FromFile($bmpPath)
    $bmp.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    Remove-Item $bmpPath -Force
    Write-Host "Converted: $pngPath"
}
